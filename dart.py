"""DART(전자공시) OpenAPI 래퍼.

- corpCode.xml : 종목코드 -> 고유번호(corp_code) 매핑
- fnlttSinglAcntAll.json : 단일회사 전체 재무제표(계정과목 + 금액)
"""

import io
import os
import re
import json
import time
import zipfile
import html as _html
import requests
import xml.etree.ElementTree as ET

API = "https://opendart.fss.or.kr/api"
# 상장사 종목코드→고유번호 매핑(사전 생성). 있으면 런타임 다운로드를 건너뛴다.
_CORP_MAP_FILE = os.path.join(os.path.dirname(os.path.abspath(__file__)), "corp_map.json")

# 캐시
_corp_cache = {}        # key -> {stock_code: {'corp_code', 'corp_name'}}
_stmt_cache = {}        # (key, corp_code, year, reprt, fs) -> list


class DartError(Exception):
    """DART API가 오류 상태를 반환했을 때."""


def get_corp_map(key):
    """상장사 종목코드 -> 고유번호 매핑. 번들 파일이 있으면 그것을 쓰고,
    없을 때만 DART에서 전체를 내려받는다(로컬 개발/최초 생성용)."""
    if key in _corp_cache:
        return _corp_cache[key]

    # 사전 생성된 매핑 파일 우선 (서버리스 콜드스타트에서 20MB 다운로드 회피)
    if os.path.exists(_CORP_MAP_FILE):
        with open(_CORP_MAP_FILE, encoding="utf-8") as f:
            m = json.load(f)
        _corp_cache[key] = m
        return m

    r = requests.get(f"{API}/corpCode.xml", params={"crtfc_key": key}, timeout=60)

    # 정상 응답은 ZIP(PK 시그니처). 오류면 JSON이 온다.
    if r.content[:2] != b"PK":
        try:
            j = json.loads(r.content.decode("utf-8", "ignore"))
            raise DartError(j.get("message", "corpCode 다운로드 실패"))
        except DartError:
            raise
        except Exception:
            raise DartError("corpCode 다운로드 실패 (API 키를 확인하세요)")

    zf = zipfile.ZipFile(io.BytesIO(r.content))
    data = zf.read(zf.namelist()[0])

    # 파일이 크므로(약 20MB) iterparse로 메모리 사용을 최소화한다.
    m = {}
    for _, el in ET.iterparse(io.BytesIO(data), events=("end",)):
        if el.tag != "list":
            continue
        stock = (el.findtext("stock_code") or "").strip()
        if stock and stock.isdigit():  # 상장사만 (종목코드 존재)
            m[stock] = {
                "corp_code": (el.findtext("corp_code") or "").strip(),
                "corp_name": (el.findtext("corp_name") or "").strip(),
            }
        el.clear()

    _corp_cache[key] = m
    return m


def get_statement(key, corp_code, year, reprt, fs_div):
    """단일회사 전체 재무제표 리스트 반환. 데이터 없으면 빈 리스트."""
    ck = (key, corp_code, year, reprt, fs_div)
    if ck in _stmt_cache:
        return _stmt_cache[ck]

    r = requests.get(
        f"{API}/fnlttSinglAcntAll.json",
        params={
            "crtfc_key": key,
            "corp_code": corp_code,
            "bsns_year": year,
            "reprt_code": reprt,
            "fs_div": fs_div,
        },
        timeout=30,
    )
    j = r.json()
    status = j.get("status")

    if status == "013":       # 조회된 데이터 없음
        _stmt_cache[ck] = []
        return []
    if status != "000":
        raise DartError(j.get("message", f"DART 오류(status={status})"))

    rows = j.get("list", [])
    _stmt_cache[ck] = rows
    return rows


# 재무제표 표시 순서: 재무상태표 → 손익 → 포괄손익 → 현금흐름 → 자본변동
_SJ_ORDER = {"BS": 0, "IS": 1, "CIS": 2, "CF": 3, "SCE": 4}

# DART의 ord 값은 account_id 알파벳순이라 실제 표시순서가 아니다.
# 표준 계정 ID 기준으로 정식 재무제표 순서를 지정한다(정확 매칭).
_CANON_EXACT = {
    # ── 재무상태표(BS): 자산 → 부채 → 자본
    "ifrs-full_CurrentAssets": 100, "ifrs-full_NoncurrentAssets": 130, "ifrs-full_Assets": 160,
    "ifrs-full_CurrentLiabilities": 200, "ifrs-full_NoncurrentLiabilities": 230, "ifrs-full_Liabilities": 260,
    "ifrs-full_IssuedCapital": 300, "dart_IssuedCapitalOfCommonStock": 301,
    "dart_IssuedCapitalOfPreferredStock": 302, "ifrs-full_SharePremium": 305,
    "ifrs-full_RetainedEarnings": 310, "dart_ElementsOfOtherStockholdersEquity": 315,
    "ifrs-full_EquityAttributableToOwnersOfParent": 320, "ifrs-full_NoncontrollingInterests": 325,
    "ifrs-full_Equity": 330, "ifrs-full_EquityAndLiabilities": 340,
    # ── 손익계산서(IS): 매출액 → 매출원가 → 매출총이익 → 판관비 → 영업이익 → … → 당기순이익 → EPS
    "ifrs-full_Revenue": 500, "ifrs-full_CostOfSales": 501, "ifrs-full_GrossProfit": 502,
    "dart_TotalSellingGeneralAdministrativeExpenses": 503,
    "dart_OperatingIncomeLoss": 504, "ifrs-full_OperatingIncomeLoss": 504,
    "dart_OtherGains": 505, "dart_OtherLosses": 506,
    "ifrs-full_OtherIncome": 507, "ifrs-full_OtherExpenseByNature": 508,
    "ifrs-full_FinanceIncome": 511, "ifrs-full_FinanceCosts": 512,
    "ifrs-full_ProfitLossBeforeTax": 520, "ifrs-full_IncomeTaxExpenseContinuingOperations": 521,
    "ifrs-full_ProfitLossFromContinuingOperations": 522, "ifrs-full_ProfitLossFromDiscontinuedOperations": 523,
    "ifrs-full_ProfitLoss": 530,
    "ifrs-full_ProfitLossAttributableToOwnersOfParent": 531,
    "ifrs-full_ProfitLossAttributableToNoncontrollingInterests": 532,
    "ifrs-full_BasicEarningsLossPerShare": 540, "ifrs-full_DilutedEarningsLossPerShare": 541,
    # ── 포괄손익계산서(CIS)
    "ifrs-full_OtherComprehensiveIncome": 602,
    "ifrs-full_ComprehensiveIncome": 650,
    "ifrs-full_ComprehensiveIncomeAttributableToOwnersOfParent": 651,
    "ifrs-full_ComprehensiveIncomeAttributableToNoncontrollingInterests": 652,
    # ── 현금흐름표(CF): 영업 → 투자 → 재무 → 기초/기말
    "ifrs-full_CashFlowsFromUsedInOperatingActivities": 700,
    "ifrs-full_CashFlowsFromUsedInInvestingActivities": 730,
    "ifrs-full_CashFlowsFromUsedInFinancingActivities": 760,
    "dart_CashAndCashEquivalentsAtBeginningOfPeriodCf": 795,
    "dart_CashAndCashEquivalentsAtEndOfPeriodCf": 796,
}
# 접두어 매칭(변형 id가 많은 항목). 구체적인 접두어를 먼저 둔다.
_CANON_PREFIX = [
    ("ifrs-full_ShareOfProfitLossOfAssociates", 510),                       # 지분법이익
    ("ifrs-full_OtherComprehensiveIncomeThatWillNotBeReclassified", 610),
    ("ifrs-full_OtherComprehensiveIncomeThatWillBeReclassified", 630),
    ("ifrs-full_GainsLossesOnRemeasurementsOfDefinedBenefitPlans", 611),
    ("ifrs-full_OtherComprehensiveIncomeNetOfTax", 615),
    ("ifrs-full_GainsLossesOnCashFlowHedges", 631),
    ("ifrs-full_GainsLossesOnExchangeDifferencesOnTranslation", 632),
    ("ifrs-full_ShareOfOtherComprehensiveIncome", 633),
    ("ifrs-full_OtherComprehensiveIncome", 620),                            # 일반(구체 접두어 뒤)
]


def _canon_rank(aid):
    if aid in _CANON_EXACT:
        return _CANON_EXACT[aid]
    for prefix, rank in _CANON_PREFIX:
        if aid.startswith(prefix):
            return rank
    return 900   # 미지정(상세/사용자정의) 계정은 각 표 뒤쪽에 배치


def _acct_sort_key(row):
    try:
        o = int(row.get("ord") or 0)
    except (TypeError, ValueError):
        o = 0
    aid = row.get("account_id") or ""
    return (_SJ_ORDER.get(row.get("sj_div"), 99), _canon_rank(aid), o)


def get_accounts(key, corp_code, year, reprt, fs_div):
    """계정과목 목록을 재무제표 표시 순서대로 반환.
    [{'sj_nm': '재무상태표', 'account_nm': '자산총계', 'sj': 0, 'ord': 7}, ...]"""
    rows = get_statement(key, corp_code, year, reprt, fs_div)
    out, seen = [], set()
    for row in sorted(rows, key=_acct_sort_key):
        nm = (row.get("account_nm") or "").strip()
        sj = (row.get("sj_nm") or "").strip()
        if nm and nm not in seen:
            seen.add(nm)
            out.append({"sj_nm": sj, "account_nm": nm, "_k": list(_acct_sort_key(row))})
    return out


# ============================================================
# 플러스알파: 사업보고서 "내용" (계정과목 외)
# ============================================================

def _get_json(ep, params, retries=3):
    """DART JSON 호출. status=101(연속호출 제한)이면 잠깐 쉬고 재시도."""
    j = {}
    for attempt in range(retries):
        r = requests.get(f"{API}/{ep}", params=params, timeout=30)
        try:
            j = r.json()
        except Exception:
            return {"status": "900", "message": "응답 파싱 실패"}
        if j.get("status") == "101" and attempt < retries - 1:
            time.sleep(0.7)
            continue
        return j
    return j


def get_major_info(key, corp_code, year, reprt, endpoint):
    """정기보고서 주요정보 단일 항목 조회 -> (list, status)."""
    j = _get_json(endpoint + ".json", {
        "crtfc_key": key, "corp_code": corp_code,
        "bsns_year": year, "reprt_code": reprt,
    })
    st = j.get("status")
    if st == "000":
        return j.get("list", []), "000"
    if st == "013":            # 조회된 데이터 없음
        return [], "013"
    return [], st or "900"


def find_report_rcept(key, corp_code, year):
    """해당 사업연도 '사업보고서'의 접수번호(rcept_no)를 찾는다. (사업보고서는 다음 해에 제출)"""
    j = _get_json("list.json", {
        "crtfc_key": key, "corp_code": corp_code,
        "bgn_de": f"{year}0101", "end_de": f"{int(year) + 1}1231",
        "pblntf_detail_ty": "A001",   # 사업보고서
        "page_count": "100",
    })
    if j.get("status") != "000":
        return None
    cands = j.get("list", [])
    for it in cands:                              # (YYYY.12) 정기 사업보고서 우선
        if f"({year}.12)" in it.get("report_nm", ""):
            return it.get("rcept_no")
    return cands[0].get("rcept_no") if cands else None


def _strip(s):
    """XML/HTML 태그 제거 후 공백 정리."""
    s = re.sub(r"<[^>]+>", " ", s)
    s = _html.unescape(s)
    return re.sub(r"\s+", " ", s).strip()


_doc_cache = {}   # (key, rcept_no) -> (xml, titles[(start, end, text)])
_RE_TOP = re.compile(r"^[IVXLC]+\.")        # 대분류: 로마숫자.
_RE_SUB = re.compile(r"^\d+(-\d+)?\.")      # 소분류: 숫자. / 숫자-숫자.


def _is_top(t):
    """대분류 여부: 로마숫자 시작 또는 【 대표이사/전문가 확인 】 같은 괄호 제목."""
    return bool(_RE_TOP.match(t)) or t.strip().startswith("【")


def _load_doc(key, rcept_no):
    """사업보고서 원본 XML과 TITLE(목차) 위치 목록을 읽어 캐시."""
    ck = (key, rcept_no)
    if ck in _doc_cache:
        return _doc_cache[ck]

    r = requests.get(f"{API}/document.xml",
                     params={"crtfc_key": key, "rcept_no": rcept_no}, timeout=90)
    if r.content[:2] != b"PK":
        _doc_cache[ck] = ("", [])
        return _doc_cache[ck]

    zf = zipfile.ZipFile(io.BytesIO(r.content))
    main = max(zf.namelist(), key=lambda n: zf.getinfo(n).file_size)  # 본문 = 최대 파일
    xml = zf.read(main).decode("utf-8", "ignore")
    titles = [(m.start(), m.end(), _strip(m.group(1)))
              for m in re.finditer(r"<TITLE[^>]*>(.*?)</TITLE>", xml, re.S)]
    titles = [(s, e, t) for (s, e, t) in titles if t]
    _doc_cache[ck] = (xml, titles)
    return _doc_cache[ck]


def get_report_toc(key, rcept_no):
    """사업보고서 전체 목차를 대분류>소분류 트리로 반환(모든 섹션 포함).
    회사의 개요 ~ 재무에 관한 사항(주석 포함) ~ 상세표 ~ 전문가의 확인까지."""
    _, titles = _load_doc(key, rcept_no)
    toc, cur = [], None
    for (s, e, t) in titles:
        if not t or t in ("목 차", "목차"):
            continue
        if _is_top(t):
            cur = {"top": t, "subs": []}
            toc.append(cur)
        elif _RE_SUB.match(t) and cur is not None:
            if t not in cur["subs"]:
                cur["subs"].append(t)
    return toc


def _find_title(titles, text):
    for i, (s, e, t) in enumerate(titles):
        if t == text:
            return i
    for i, (s, e, t) in enumerate(titles):
        if text and (text in t or t in text):
            return i
    return None


def _section_slice(key, rcept_no, top, sub=None):
    """대분류(top) 범위 안에서 소분류(sub)에 해당하는 원본 XML 조각을 반환."""
    xml, titles = _load_doc(key, rcept_no)
    if not titles:
        return ""
    ti = _find_title(titles, top)
    if ti is None:
        return ""

    top_start = titles[ti][1]
    top_end = len(xml)                       # 다음 대분류 전까지가 이 대분류 범위
    for j in range(ti + 1, len(titles)):
        if _is_top(titles[j][2]):
            top_end = titles[j][0]
            break

    if not sub or sub == top:
        return xml[top_start:top_end]

    si = None                                # 대분류 범위 안에서 소분류 찾기
    for j in range(ti + 1, len(titles)):
        if titles[j][0] >= top_end:
            break
        if titles[j][2] == sub or sub in titles[j][2]:
            si = j
            break
    if si is None:
        return ""

    sub_start = titles[si][1]
    sub_end = top_end
    for j in range(si + 1, len(titles)):     # 다음 소분류(또는 대분류 끝)까지
        if titles[j][0] >= top_end:
            break
        if _RE_SUB.match(titles[j][2]):
            sub_end = titles[j][0]
            break
    return xml[sub_start:sub_end]


def get_section_text(key, rcept_no, top, sub=None):
    """소분류 텍스트(태그 제거)."""
    return _strip(_section_slice(key, rcept_no, top, sub))


def _parse_table(tbl_xml):
    """<TABLE> 조각을 2차원 배열(행 x 셀)로 변환."""
    rows = []
    for tr in re.findall(r"<TR[^>]*>(.*?)</TR>", tbl_xml, re.S):
        cells = re.findall(r"<(?:TD|TE|TH|TU)[^>]*>(.*?)</(?:TD|TE|TH|TU)>", tr, re.S)
        row = [_strip(catch) for catch in cells]
        if any(row):
            rows.append(row)
    return rows


def _add_text_blocks(blocks, frag):
    """XML 조각에서 <P> 문단들을 텍스트 블록으로 추가."""
    ps = re.findall(r"<P[^>]*>(.*?)</P>", frag, re.S)
    for p in (ps if ps else [frag]):
        t = _strip(p)
        if t:
            blocks.append(("text", t))


def get_section_blocks(key, rcept_no, top, sub=None):
    """소분류 내용을 문단/표 블록 리스트로 반환.
    [('text', '문단...'), ('table', [[셀,셀],[...]]), ...]"""
    frag = _section_slice(key, rcept_no, top, sub)
    if not frag:
        return []
    blocks, pos = [], 0
    for m in re.finditer(r"<TABLE[^>]*>.*?</TABLE>", frag, re.S):
        _add_text_blocks(blocks, frag[pos:m.start()])
        rows = _parse_table(m.group(0))
        if rows:
            blocks.append(("table", rows))
        pos = m.end()
    _add_text_blocks(blocks, frag[pos:])
    return blocks


def get_auditor(key, rcept_no):
    """외부감사인 정보 추출. 'V.회계감사인의 감사의견 등 > 1.외부감사에 관한 사항'의
    감사인/감사의견 표에서 사업연도별 [{term, auditor, opinion}] 반환."""
    blocks = get_section_blocks(key, rcept_no,
                                "V. 회계감사인의 감사의견 등", "1. 외부감사에 관한 사항")
    for kind, payload in blocks:
        if kind != "table" or len(payload) < 2:
            continue
        header = payload[0]
        c_term = c_aud = c_op = None
        for i, h in enumerate(header):
            h = h or ""
            if c_term is None and ("사업연도" in h or "기수" in h):
                c_term = i
            if c_aud is None and "감사인" in h:
                c_aud = i
            if c_op is None and "감사의견" in h:
                c_op = i
        if c_aud is None or c_op is None:
            continue

        def cell(row, i):
            return (row[i].strip() if (i is not None and i < len(row) and row[i]) else "")

        out = []
        for row in payload[1:]:
            aud = cell(row, c_aud)
            if aud:
                out.append({"term": cell(row, c_term), "auditor": aud,
                            "opinion": cell(row, c_op)})
        if out:
            return out
    return []

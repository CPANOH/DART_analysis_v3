"""DART(전자공시) OpenAPI 래퍼.

- corpCode.xml : 종목코드 -> 고유번호(corp_code) 매핑
- fnlttSinglAcntAll.json : 단일회사 전체 재무제표(계정과목 + 금액)
"""

import io
import re
import json
import time
import zipfile
import html as _html
import requests
import xml.etree.ElementTree as ET

API = "https://opendart.fss.or.kr/api"

# 캐시
_corp_cache = {}        # key -> {stock_code: {'corp_code', 'corp_name'}}
_stmt_cache = {}        # (key, corp_code, year, reprt, fs) -> list


class DartError(Exception):
    """DART API가 오류 상태를 반환했을 때."""


def get_corp_map(key):
    """DART 고유번호 전체를 내려받아 종목코드 -> 정보 딕셔너리로 반환."""
    if key in _corp_cache:
        return _corp_cache[key]

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


def get_accounts(key, corp_code, year, reprt, fs_div):
    """계정과목 목록만 추출. [{'sj_nm': '재무상태표', 'account_nm': '자산총계'}, ...]"""
    rows = get_statement(key, corp_code, year, reprt, fs_div)
    out, seen = [], set()
    for row in rows:
        nm = (row.get("account_nm") or "").strip()
        sj = (row.get("sj_nm") or "").strip()
        if nm and nm not in seen:
            seen.add(nm)
            out.append({"sj_nm": sj, "account_nm": nm})
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
    """사업보고서 목차를 대분류>소분류 트리로 반환.
    재무제표(계정과목으로 이미 다룸)와 상세표는 제외. [{'top':..., 'subs':[...]}]"""
    _, titles = _load_doc(key, rcept_no)
    toc, cur = [], None
    for (s, e, t) in titles:
        if _RE_TOP.match(t):
            skip = ("재무에 관한" in t) or ("상세표" in t)
            cur = None if skip else {"top": t, "subs": []}
            if cur is not None:
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


def get_section_text(key, rcept_no, top, sub=None):
    """대분류(top) 범위 안에서 소분류(sub) 텍스트를 추출. sub가 없으면 대분류 전체."""
    xml, titles = _load_doc(key, rcept_no)
    if not titles:
        return ""
    ti = _find_title(titles, top)
    if ti is None:
        return ""

    top_start = titles[ti][1]
    top_end = len(xml)                       # 다음 대분류 전까지가 이 대분류 범위
    for j in range(ti + 1, len(titles)):
        if _RE_TOP.match(titles[j][2]):
            top_end = titles[j][0]
            break

    if not sub or sub == top:
        return _strip(xml[top_start:top_end])

    # 대분류 범위 안에서 소분류 찾기
    si = None
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
    return _strip(xml[sub_start:sub_end])

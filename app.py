"""재무분석기 - 로컬 웹앱

1차 필터: 산업군(네이버 증권)  ->  2차 필터: 회사명(네이버+DART)  ->  3차 필터: 계정과목(DART)
선택한 자료를 엑셀(.xlsx)로 다운로드.
"""

import io
import os
import re
import time
from collections import OrderedDict
from datetime import datetime

from flask import Flask, request, jsonify, send_file, render_template

import naver
import dart
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment, PatternFill
from openpyxl.utils import get_column_letter

BASE_DIR = os.path.dirname(os.path.abspath(__file__))


def _load_dotenv():
    """로컬 .env 파일이 있으면 환경변수로 읽어들인다(외부 의존성 없이)."""
    path = os.path.join(BASE_DIR, ".env")
    if not os.path.exists(path):
        return
    with open(path, "r", encoding="utf-8") as f:
        for line in f:
            line = line.strip()
            if not line or line.startswith("#") or "=" not in line:
                continue
            k, v = line.split("=", 1)
            os.environ.setdefault(k.strip(), v.strip().strip('"').strip("'"))


_load_dotenv()

app = Flask(__name__, template_folder=os.path.join(BASE_DIR, "templates"))


def resolve_key(client_key):
    """클라이언트가 보낸 키를 우선하고, 없으면 서버 환경변수(DART_API_KEY)를 사용."""
    return (client_key or "").strip() or os.environ.get("DART_API_KEY", "").strip()

REPRT_NAMES = {
    "11011": "사업보고서(연간)",
    "11012": "반기보고서",
    "11013": "1분기보고서",
    "11014": "3분기보고서",
}
FS_NAMES = {"CFS": "연결재무제표", "OFS": "별도재무제표"}

# --- 플러스알파: 사업보고서 "내용" 카탈로그 -------------------------------

# 정기보고서 주요정보 (구조화 데이터): 그룹 > 세부항목(endpoint)
MAJOR_GROUPS = [
    {"group": "배당", "items": [
        {"id": "alotMatter", "name": "배당에 관한 사항"},
    ]},
    {"group": "주식·자본", "items": [
        {"id": "irdsSttus", "name": "증자(감자) 현황"},
        {"id": "tesstkAcqsDspsSttus", "name": "자기주식 취득·처분"},
    ]},
    {"group": "주주", "items": [
        {"id": "hyslrSttus", "name": "최대주주 현황"},
        {"id": "hyslrChgSttus", "name": "최대주주 변동현황"},
        {"id": "mrhlSttus", "name": "소액주주 현황"},
    ]},
    {"group": "임직원·보수", "items": [
        {"id": "exctvSttus", "name": "임원 현황"},
        {"id": "empSttus", "name": "직원 현황"},
        {"id": "hmvAuditAllSttus", "name": "이사·감사 전체 보수현황"},
        {"id": "hmvAuditIndvdlBySttus", "name": "개인별 보수(5억 이상)"},
    ]},
    {"group": "투자·채무", "items": [
        {"id": "otrCprInvstmntSttus", "name": "타법인 출자현황"},
        {"id": "cprndNrdmpBlce", "name": "회사채 미상환 잔액"},
        {"id": "srtpdPsndbtNrdmpBlce", "name": "단기사채 미상환 잔액"},
        {"id": "entrprsBilScritsNrdmpBlce", "name": "기업어음 미상환 잔액"},
        {"id": "detScritsIsuAcmslt", "name": "채무증권 발행실적"},
        {"id": "pssrpCptalUseDtls", "name": "공모자금 사용내역"},
        {"id": "prvsrpCptalUseDtls", "name": "사모자금 사용내역"},
    ]},
]
MAJOR_MAP = {it["id"]: it["name"] for g in MAJOR_GROUPS for it in g["items"]}

# DART 필드코드 -> 한글 라벨 (주요정보 시트 헤더용; 없으면 코드 그대로)
FIELD_LABELS = {
    "se": "구분", "stock_knd": "주식종류", "thstrm": "당기", "frmtrm": "전기", "lwfr": "전전기",
    "nm": "성명", "sexdstn": "성별", "birth_ym": "출생년월", "ofcps": "직위",
    "rgist_exctv_at": "등기임원여부", "fte_at": "상근여부", "chrg_job": "담당업무",
    "main_career": "주요경력", "mxmm_shrholdr_relate": "최대주주관계",
    "hffc_pd": "재직기간", "tenure_end_on": "임기만료일", "relate": "관계",
    "bsis_posesn_stock_co": "기초주식수", "bsis_posesn_stock_qota_rt": "기초지분율(%)",
    "trmend_posesn_stock_co": "기말주식수", "trmend_posesn_stock_qota_rt": "기말지분율(%)",
    "fo_bbm": "사업부문", "rgllbr_co": "정규직수", "cnttk_co": "계약직수", "sm": "합계",
    "avrg_cnwk_sdytrn": "평균근속연수", "fyer_salary_totamt": "연간급여총액",
    "jan_salary_am": "1인평균급여", "shrholdr_co": "소액주주수", "shrholdr_tot_co": "전체주주수",
    "shrholdr_rate": "소액주주비율(%)", "hold_stock_co": "소액주주보유주식수",
    "stock_tot_co": "총발행주식수", "hold_stock_rate": "소액주주지분율(%)",
    "acqs_mth1": "취득방법1", "acqs_mth2": "취득방법2", "acqs_mth3": "취득방법3",
    "bsis_qy": "기초수량", "change_qy_acqs": "취득수량", "change_qy_dsps": "처분수량",
    "change_qy_incnr": "소각수량", "trmend_qy": "기말수량",
    "isu_dcrs_de": "발행일", "isu_dcrs_stle": "발행형태", "isu_dcrs_stock_knd": "주식종류",
    "isu_dcrs_qy": "수량", "isu_dcrs_mstvdv_fval_amount": "액면가",
    "isu_dcrs_mstvdv_amount": "발행가", "rm": "비고", "stlm_dt": "결산기준일",
}
# 시트에서 숨길 반복 메타 필드
SKIP_FIELDS = {"rcept_no", "corp_cls", "corp_code", "corp_name"}
CELL_MAX = 32000   # 엑셀 셀 최대 문자수(32,767) 안전 여유


def _num(s):
    """DART 금액 문자열 -> 숫자(가능하면 int). 실패 시 원본 문자열 또는 None."""
    if s is None:
        return None
    s = str(s).strip().replace(",", "")
    if s in ("", "-"):
        return None
    try:
        return int(s)
    except ValueError:
        try:
            return float(s)
        except ValueError:
            return s


# ------------------------------------------------------------------ 라우트

@app.route("/")
def index():
    return render_template("index.html")


@app.route("/api/config")
def api_config():
    # 서버에 키가 설정돼 있으면 프론트에서 키 입력을 선택사항으로 안내.
    return jsonify({
        "has_server_key": bool(os.environ.get("DART_API_KEY", "").strip()),
        "major_groups": MAJOR_GROUPS,
    })


@app.route("/api/toc")
def api_toc():
    """선택 회사의 사업보고서 목차(대분류>소분류)를 반환."""
    key = resolve_key(request.args.get("key"))
    corp_code = request.args.get("corp_code", "")
    year = request.args.get("year", "")
    if not key:
        return jsonify({"ok": False, "error": "DART API 키가 필요합니다."}), 400
    if not corp_code:
        return jsonify({"ok": False, "error": "회사를 먼저 선택하세요."}), 400
    try:
        rcept = dart.find_report_rcept(key, corp_code, year)
        if not rcept:
            return jsonify({"ok": True, "toc": [], "msg": "해당 연도 사업보고서를 찾지 못했습니다."})
        return jsonify({"ok": True, "toc": dart.get_report_toc(key, rcept)})
    except dart.DartError as e:
        return jsonify({"ok": False, "error": f"DART: {e}"}), 400
    except Exception as e:
        return jsonify({"ok": False, "error": f"목차 로드 실패: {e}"}), 500


@app.route("/api/industries")
def api_industries():
    try:
        return jsonify({"ok": True, "industries": naver.get_industries()})
    except Exception as e:
        return jsonify({"ok": False, "error": f"산업군 로드 실패: {e}"}), 500


@app.route("/api/companies")
def api_companies():
    no = request.args.get("no", "")
    key = resolve_key(request.args.get("key"))
    if not no:
        return jsonify({"ok": False, "error": "업종(no)이 필요합니다."}), 400
    try:
        companies = naver.get_companies(no)
        corp_map = dart.get_corp_map(key) if key else {}
        out = []
        for c in companies:
            info = corp_map.get(c["stock_code"])
            out.append({
                "name": c["name"],
                "stock_code": c["stock_code"],
                "corp_code": info["corp_code"] if info else None,
                "in_dart": bool(info),
            })
        return jsonify({"ok": True, "companies": out})
    except dart.DartError as e:
        return jsonify({"ok": False, "error": f"DART: {e}"}), 400
    except Exception as e:
        return jsonify({"ok": False, "error": f"회사 로드 실패: {e}"}), 500


@app.route("/api/accounts")
def api_accounts():
    key = resolve_key(request.args.get("key"))
    corp_codes = [c for c in request.args.get("corp_codes", "").split(",") if c]
    year = request.args.get("year", "")
    reprt = request.args.get("reprt", "11011")
    fs = request.args.get("fs", "CFS")

    if not key:
        return jsonify({"ok": False, "error": "DART API 키가 필요합니다."}), 400
    if not corp_codes:
        return jsonify({"ok": False, "error": "회사를 먼저 선택하세요."}), 400

    try:
        merged, seen = [], set()
        for cc in corp_codes:
            for a in dart.get_accounts(key, cc, year, reprt, fs):
                if a["account_nm"] not in seen:
                    seen.add(a["account_nm"])
                    merged.append(a)
        return jsonify({"ok": True, "accounts": merged})
    except dart.DartError as e:
        return jsonify({"ok": False, "error": f"DART: {e}"}), 400
    except Exception as e:
        return jsonify({"ok": False, "error": f"계정과목 로드 실패: {e}"}), 500


HEADER_FILL = PatternFill("solid", fgColor="1F4E78")
HEADER_FONT = Font(color="FFFFFF", bold=True)


def _safe_sheet_title(wb, name):
    """엑셀 시트명 제약(31자, 특수문자 금지, 중복불가) 처리."""
    name = re.sub(r"[\[\]:*?/\\]", " ", name).strip()[:31] or "시트"
    base, i, existing = name, 1, set(wb.sheetnames)
    while name in existing:
        i += 1
        suffix = f"_{i}"
        name = base[:31 - len(suffix)] + suffix
    return name


def _write_header_row(ws, row, headers):
    for col, h in enumerate(headers, start=1):
        cell = ws.cell(row=row, column=col, value=h)
        cell.fill = HEADER_FILL
        cell.font = HEADER_FONT
        cell.alignment = Alignment(horizontal="center")


def _build_major_sheet(ws, key, companies, years, reprt, endpoint):
    """정기보고서 주요정보 1개 항목을 회사×연도별로 모아 시트 작성."""
    all_rows, cols = [], []   # all_rows: [(company_name, year, row_dict)]
    for c in companies:
        for y in years:
            rows, _ = dart.get_major_info(key, c["corp_code"], y, reprt, endpoint)
            for r in rows:
                all_rows.append((c["name"], y, r))
                if not cols:
                    cols = [k for k in r.keys() if k not in SKIP_FIELDS]
            time.sleep(0.12)   # 연속 호출 제한 완화

    headers = ["회사명", "연도"] + [FIELD_LABELS.get(k, k) for k in cols]
    _write_header_row(ws, 1, headers)
    for ri, (cname, y, r) in enumerate(all_rows, start=2):
        ws.cell(row=ri, column=1, value=cname)
        ws.cell(row=ri, column=2, value=y)
        for ci, k in enumerate(cols, start=3):
            ws.cell(row=ri, column=ci, value=r.get(k))
    if not all_rows:
        ws.cell(row=2, column=1, value="(조회된 데이터 없음)")

    ws.column_dimensions["A"].width = 18
    ws.column_dimensions["B"].width = 8
    for ci in range(3, len(headers) + 1):
        ws.column_dimensions[get_column_letter(ci)].width = 16
    ws.freeze_panes = "A2"


def _short_top(top):
    """대분류 제목에서 로마숫자 접두어 제거 ('II. 사업의 내용' -> '사업의 내용')."""
    return re.sub(r"^[IVXLC]+\.\s*", "", top).strip() or top


def _build_section_sheet(ws, key, companies, years, top, subs):
    """한 대분류(top)의 선택된 소분류(subs)들을 회사×연도별로 시트에 작성.
    행: 회사명 | 연도 | 소분류 | 내용.  subs가 비면 대분류 전체."""
    _write_header_row(ws, 1, ["회사명", "연도", "소분류", "내용"])
    targets = subs if subs else [top]
    r = 2
    for c in companies:
        for y in years:
            rcept = dart.find_report_rcept(key, c["corp_code"], y)
            for sub in targets:
                text = dart.get_section_text(key, rcept, top, sub) if rcept else ""
                if not text:
                    text = "(내용 없음)"
                label = "전체" if sub == top else sub
                chunks = [text[i:i + CELL_MAX] for i in range(0, len(text), CELL_MAX)][:5] or [""]
                for j, ch in enumerate(chunks):
                    ws.cell(row=r, column=1, value=c["name"] if j == 0 else "")
                    ws.cell(row=r, column=2, value=y if j == 0 else "")
                    ws.cell(row=r, column=3, value=label if j == 0 else "(이어서)")
                    cell = ws.cell(row=r, column=4, value=ch)
                    cell.alignment = Alignment(wrap_text=True, vertical="top")
                    r += 1
    ws.column_dimensions["A"].width = 16
    ws.column_dimensions["B"].width = 8
    ws.column_dimensions["C"].width = 26
    ws.column_dimensions["D"].width = 100
    ws.freeze_panes = "A2"


def _build_accounts_sheet(ws, key, companies, years, reprt, fs, accounts, industry):
    """계정과목 비교 시트. 열 = 회사(병합) × 연도, 행 = 계정과목."""
    per = len(years)
    col_specs = []   # [(company_name, year, {account_nm: amount})]  회사-major 순서
    for c in companies:
        for y in years:
            rows = dart.get_statement(key, c["corp_code"], y, reprt, fs)
            amap = {}
            for r in rows:
                nm = (r.get("account_nm") or "").strip()
                if nm and nm not in amap:
                    amap[nm] = _num(r.get("thstrm_amount"))
            col_specs.append((c["name"], y, amap))

    meta = [
        "재무분석기 · CPA KOOK",
        f"산업군: {industry}",
        f"연도: {', '.join(str(y) for y in years)}   보고서: {REPRT_NAMES.get(reprt, reprt)}   구분: {FS_NAMES.get(fs, fs)}",
        f"생성일시: {datetime.now():%Y-%m-%d %H:%M}   |   작성: CPA KOOK",
    ]
    for i, line in enumerate(meta, start=1):
        ws.cell(row=i, column=1, value=line).font = Font(bold=(i == 1), size=13 if i == 1 else 10)

    h1 = len(meta) + 2   # 회사명 헤더(병합)
    h2 = h1 + 1          # 연도 헤더

    # 계정과목 헤더(2행 병합)
    ws.merge_cells(start_row=h1, start_column=1, end_row=h2, end_column=1)
    cell = ws.cell(row=h1, column=1, value="계정과목")
    cell.fill = HEADER_FILL; cell.font = HEADER_FONT
    cell.alignment = Alignment(horizontal="center", vertical="center")

    # 회사(병합) + 연도 헤더
    for ci, c in enumerate(companies):
        start = 2 + ci * per
        for yi, y in enumerate(years):
            yc = ws.cell(row=h2, column=start + yi, value=y)
            yc.fill = HEADER_FILL; yc.font = HEADER_FONT
            yc.alignment = Alignment(horizontal="center")
        if per > 1:
            ws.merge_cells(start_row=h1, start_column=start, end_row=h1, end_column=start + per - 1)
        nc = ws.cell(row=h1, column=start, value=c["name"])
        nc.fill = HEADER_FILL; nc.font = HEADER_FONT
        nc.alignment = Alignment(horizontal="center")

    # 본문
    for i, acc in enumerate(accounts):
        r = h2 + 1 + i
        ws.cell(row=r, column=1, value=acc)
        for j, (_, _, amap) in enumerate(col_specs, start=2):
            v = amap.get(acc)
            cell = ws.cell(row=r, column=j, value=v)
            if isinstance(v, (int, float)):
                cell.number_format = "#,##0"

    ws.column_dimensions["A"].width = 28
    for j in range(2, len(col_specs) + 2):
        ws.column_dimensions[get_column_letter(j)].width = 16
    ws.freeze_panes = ws.cell(row=h2 + 1, column=2)


@app.route("/api/export", methods=["POST"])
def api_export():
    data = request.get_json(force=True)
    key = resolve_key(data.get("key"))
    # 다중 연도(최대 5). 과거 호환: year 단일값도 허용.
    years = data.get("years") or ([data.get("year")] if data.get("year") else [])
    years = [str(y) for y in years if str(y).strip()][:5]
    reprt = data.get("reprt", "11011")
    fs = data.get("fs", "CFS")
    industry = data.get("industry", "")
    companies = data.get("companies", [])   # [{name, corp_code}]
    accounts = data.get("accounts", [])     # [account_nm, ...]
    # 플러스알파(선택): 사업보고서 내용
    major = [m for m in data.get("major", []) if m in MAJOR_MAP]   # 주요정보 endpoint id
    # 본문 섹션: [{top, sub}] -> 대분류별로 소분류 묶기
    sec_by_top = OrderedDict()
    for it in data.get("sections", []):
        top = (it or {}).get("top")
        sub = (it or {}).get("sub")
        if not top:
            continue
        sec_by_top.setdefault(top, [])
        if sub and sub != top and sub not in sec_by_top[top]:
            sec_by_top[top].append(sub)
    sections = sec_by_top

    if not (key and companies):
        return jsonify({"ok": False, "error": "API 키와 회사를 선택하세요."}), 400
    if not years:
        return jsonify({"ok": False, "error": "연도를 1개 이상 선택하세요."}), 400
    if not (accounts or major or sections):
        return jsonify({"ok": False, "error": "계정과목 또는 사업보고서 내용을 하나 이상 선택하세요."}), 400

    try:
        wb = Workbook()
        first_used = False

        # ---- 1) 계정과목 시트 (선택 시) --------------------------------
        if accounts:
            ws = wb.active
            ws.title = "재무분석"
            first_used = True
            _build_accounts_sheet(ws, key, companies, years, reprt, fs, accounts, industry)

        # ---- 2) 정기보고서 주요정보 시트들 -----------------------------
        for mid in major:
            ws = wb.active if not first_used else wb.create_sheet()
            ws.title = _safe_sheet_title(wb, MAJOR_MAP[mid])
            first_used = True
            _build_major_sheet(ws, key, companies, years, reprt, mid)

        # ---- 3) 사업보고서 본문 섹션 시트들 (대분류별 시트) -------------
        for top, subs in sections.items():
            ws = wb.active if not first_used else wb.create_sheet()
            ws.title = _safe_sheet_title(wb, _short_top(top))
            first_used = True
            _build_section_sheet(ws, key, companies, years, top, subs)

    except dart.DartError as e:
        return jsonify({"ok": False, "error": f"DART: {e}"}), 400
    except Exception as e:
        return jsonify({"ok": False, "error": f"데이터 수집 실패: {e}"}), 500

    bio = io.BytesIO()
    wb.save(bio)
    bio.seek(0)
    yr_tag = "-".join(str(y) for y in years)
    fname = f"재무분석_{industry}_{yr_tag}_{datetime.now():%Y%m%d_%H%M%S}.xlsx"
    return send_file(
        bio,
        as_attachment=True,
        download_name=fname,
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )


if __name__ == "__main__":
    import webbrowser, threading
    url = "http://127.0.0.1:5000"
    threading.Timer(1.2, lambda: webbrowser.open(url)).start()
    print(f"\n  재무분석기 실행 중 -> {url}\n  (종료하려면 이 창에서 Ctrl+C)\n")
    app.run(host="127.0.0.1", port=5000, debug=False)

"""재무분석기 - 로컬 웹앱

1차 필터: 산업군(네이버 증권)  ->  2차 필터: 회사명(네이버+DART)  ->  3차 필터: 계정과목(DART)
선택한 자료를 엑셀(.xlsx)로 다운로드.
"""

import io
import os
from datetime import datetime

from flask import Flask, request, jsonify, send_file, render_template

import naver
import dart
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment, PatternFill
from openpyxl.utils import get_column_letter

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
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
    return jsonify({"has_server_key": bool(os.environ.get("DART_API_KEY", "").strip())})


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


@app.route("/api/export", methods=["POST"])
def api_export():
    data = request.get_json(force=True)
    key = resolve_key(data.get("key"))
    year = data.get("year", "")
    reprt = data.get("reprt", "11011")
    fs = data.get("fs", "CFS")
    industry = data.get("industry", "")
    companies = data.get("companies", [])   # [{name, corp_code}]
    accounts = data.get("accounts", [])     # [account_nm, ...]

    if not (key and companies and accounts):
        return jsonify({"ok": False, "error": "API 키/회사/계정과목이 모두 필요합니다."}), 400

    try:
        # 회사별 재무제표 -> {account_nm: 당기금액}
        col_data = []  # [(company_name, {account_nm: amount})]
        for c in companies:
            rows = dart.get_statement(key, c["corp_code"], year, reprt, fs)
            amap = {}
            for r in rows:
                nm = (r.get("account_nm") or "").strip()
                if nm and nm not in amap:
                    amap[nm] = _num(r.get("thstrm_amount"))
            col_data.append((c["name"], amap))
    except dart.DartError as e:
        return jsonify({"ok": False, "error": f"DART: {e}"}), 400
    except Exception as e:
        return jsonify({"ok": False, "error": f"데이터 수집 실패: {e}"}), 500

    # ---- 엑셀 생성 -------------------------------------------------
    wb = Workbook()
    ws = wb.active
    ws.title = "재무분석"

    header_fill = PatternFill("solid", fgColor="1F4E78")
    header_font = Font(color="FFFFFF", bold=True)

    # 상단 메타 정보
    meta = [
        f"산업군: {industry}",
        f"사업연도: {year}   보고서: {REPRT_NAMES.get(reprt, reprt)}   구분: {FS_NAMES.get(fs, fs)}",
        f"생성일시: {datetime.now():%Y-%m-%d %H:%M}",
    ]
    for i, line in enumerate(meta, start=1):
        ws.cell(row=i, column=1, value=line).font = Font(bold=(i == 1), size=12 if i == 1 else 10)

    head_row = len(meta) + 2
    # 헤더: 계정과목 | 회사1 | 회사2 | ...
    headers = ["계정과목"] + [name for name, _ in col_data]
    for col, h in enumerate(headers, start=1):
        cell = ws.cell(row=head_row, column=col, value=h)
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = Alignment(horizontal="center")

    # 본문
    for i, acc in enumerate(accounts):
        r = head_row + 1 + i
        ws.cell(row=r, column=1, value=acc)
        for j, (_, amap) in enumerate(col_data, start=2):
            v = amap.get(acc)
            cell = ws.cell(row=r, column=j, value=v)
            if isinstance(v, (int, float)):
                cell.number_format = "#,##0"

    # 열 너비
    ws.column_dimensions["A"].width = 28
    for col in range(2, len(headers) + 1):
        ws.column_dimensions[get_column_letter(col)].width = 18
    ws.freeze_panes = ws.cell(row=head_row + 1, column=2)

    bio = io.BytesIO()
    wb.save(bio)
    bio.seek(0)
    fname = f"재무분석_{industry}_{year}_{datetime.now():%Y%m%d_%H%M%S}.xlsx"
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

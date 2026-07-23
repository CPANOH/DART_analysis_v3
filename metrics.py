"""분석지표 엔진 — 표준 계정ID 기준으로 재무지표를 자동 계산한다(회사명이 달라도 매칭).

반도체 분석 세트: 매출/원가/이익, 유형자산·CAPEX, 재고, 현금흐름 관련 핵심 지표 + 증감/비율.
감가상각비·재고평가충당금 등 '주석에만 있는' 항목은 2단계에서 추가한다.
"""

import dart


def _num(s):
    if s is None:
        return None
    s = str(s).strip().replace(",", "")
    if s in ("", "-"):
        return None
    neg = s.startswith("(") and s.endswith(")")
    if neg:
        s = s[1:-1]
    try:
        v = float(s)
    except ValueError:
        return None
    return -v if neg else v


def _find(rows, ids=(), sj=None, name_contains=(), field="thstrm_amount"):
    """account_id 우선, 없으면 (해당 재무제표 내) 계정명 부분일치로 값 찾기.
    field='frmtrm_amount'면 전기값을 반환."""
    for r in rows:
        if sj and r.get("sj_div") != sj:
            continue
        if r.get("account_id") in ids:
            return _num(r.get(field))
    if name_contains:
        for r in rows:
            if sj and r.get("sj_div") != sj:
                continue
            nm = r.get("account_nm") or ""
            if any(k in nm for k in name_contains):
                return _num(r.get(field))
    return None


def _base_values(rows):
    """한 회사·한 연도의 재무제표 rows에서 기초 계정값을 추출."""
    # 손익 항목은 회사에 따라 손익계산서(IS) 또는 포괄손익계산서(CIS)로 보고되므로
    # sj_div로 제한하지 않고 표준 account_id로 매칭한다.
    rev = _find(rows, {"ifrs-full_Revenue"}, None, ("매출액", "영업수익"))
    cogs = _find(rows, {"ifrs-full_CostOfSales"}, None, ("매출원가",))
    gp = _find(rows, {"ifrs-full_GrossProfit"}, None, ("매출총이익",))
    if gp is None and rev is not None and cogs is not None:
        gp = rev - cogs
    sga = _find(rows, {"dart_TotalSellingGeneralAdministrativeExpenses"}, None,
                ("판매비와관리비", "판매비와 관리비"))
    op = _find(rows, {"dart_OperatingIncomeLoss", "ifrs-full_OperatingIncomeLoss"}, None,
               ("영업이익",))
    ni = _find(rows, {"ifrs-full_ProfitLoss"}, None, ("당기순이익",))
    ppe = _find(rows, {"ifrs-full_PropertyPlantAndEquipment"}, "BS", ("유형자산",))
    ppe_prev = _find(rows, {"ifrs-full_PropertyPlantAndEquipment"}, "BS", ("유형자산",), field="frmtrm_amount")
    inv = _find(rows, {"ifrs-full_Inventories"}, "BS", ("재고자산",))
    ocf = _find(rows, {"ifrs-full_CashFlowsFromUsedInOperatingActivities"}, "CF",
                ("영업활동 현금흐름", "영업활동현금흐름", "영업활동으로"))
    capex = _find(rows, {"ifrs-full_PurchaseOfPropertyPlantAndEquipmentClassifiedAsInvestingActivities"},
                  "CF", ("유형자산의 취득", "유형자산 취득", "유형자산의취득"))
    capex_int = _find(rows, {"ifrs-full_PurchaseOfIntangibleAssetsClassifiedAsInvestingActivities"},
                      "CF", ("무형자산의 취득", "무형자산 취득", "무형자산의취득"))
    # 총 CAPEX = 유형 + 무형(있으면). 유형이 없으면 총도 None.
    capex_total = None if capex is None else capex + (capex_int or 0)
    ppe_chg = (ppe - ppe_prev) if (ppe is not None and ppe_prev is not None) else None
    return {"rev": rev, "cogs": cogs, "gp": gp, "sga": sga, "op": op, "ni": ni,
            "ppe": ppe, "ppe_prev": ppe_prev, "ppe_chg": ppe_chg, "inv": inv, "ocf": ocf,
            "capex": capex, "capex_int": capex_int, "capex_total": capex_total}


def _div(a, b):
    if a is None or not b:
        return None
    return a / b


# 지표 정의: (표시명, 종류, 계산함수)  종류: amount/pct/ratio/days
# b = 기초값 dict
METRIC_DEFS = [
    ("매출액", "amount", lambda b: b["rev"]),
    ("매출원가", "amount", lambda b: b["cogs"]),
    ("매출총이익", "amount", lambda b: b["gp"]),
    ("판매비와관리비", "amount", lambda b: b["sga"]),
    ("영업이익", "amount", lambda b: b["op"]),
    ("당기순이익", "amount", lambda b: b["ni"]),
    ("유형자산", "amount", lambda b: b["ppe"]),
    ("재고자산", "amount", lambda b: b["inv"]),
    ("영업활동현금흐름", "amount", lambda b: b["ocf"]),
    ("유형자산 취득(CAPEX)", "amount", lambda b: b["capex"]),
    ("무형자산 취득", "amount", lambda b: b["capex_int"]),
    ("총 CAPEX(유형+무형)", "amount", lambda b: b["capex_total"]),
    ("단순 FCF(영업CF−총CAPEX)", "amount",
     lambda b: (b["ocf"] - b["capex_total"]) if (b["ocf"] is not None and b["capex_total"] is not None) else None),
    ("__sep__", "sep", None),
    ("매출원가율", "pct", lambda b: _div(b["cogs"], b["rev"])),
    ("매출총이익률", "pct", lambda b: _div(b["gp"], b["rev"])),
    ("영업이익률", "pct", lambda b: _div(b["op"], b["rev"])),
    ("순이익률", "pct", lambda b: _div(b["ni"], b["rev"])),
    ("총 CAPEX / 매출액", "pct", lambda b: _div(b["capex_total"], b["rev"])),
    ("매출액 / 유형자산(회전)", "ratio", lambda b: _div(b["rev"], b["ppe"])),
    ("재고자산회전율(매출원가÷재고)", "ratio", lambda b: _div(b["cogs"], b["inv"])),
    ("재고자산회전기간(일)", "days",
     lambda b: _div(365 * b["inv"], b["cogs"]) if (b["inv"] is not None and b["cogs"]) else None),
    ("영업CF / 영업이익", "ratio", lambda b: _div(b["ocf"], b["op"])),
]


# --- 주석 기반: 매출원가에 포함된 감가상각비·무형자산상각비(추정) ---

def _first_num(cells):
    """행에서 첫 숫자값(=당기). 구형 보고서는 한 표에 당기·전기가 나란히 있어 첫 값이 당기."""
    for c in cells:
        v = _num(c)
        if v is not None:
            return v
    return None


def _table_da(table):
    """표에서 감가상각비+무형자산상각비 행의 당기값 합계(백만원)."""
    total, found = 0.0, False
    for row in table[1:]:
        lab = row[0] or ""
        if ("감가상각" in lab) or ("무형자산상각" in lab):
            v = _first_num(row[1:])
            if v is not None:
                total += v
                found = True
    return total if found else None


def _da_from_first_table(key, rcept, top, sub):
    """주석 소분류의 (당기) 표에서 D&A 합계(백만원)."""
    for kind, payload in dart.get_section_blocks(key, rcept, top, sub):
        if kind == "table" and len(payload) >= 2:
            da = _table_da(payload)
            if da is not None:
                return da
    return None


def _find_note(subs, keyword, consolidated):
    cands = [s for s in subs if keyword in s]
    if not cands:
        return None
    pref = [s for s in cands if (("(연결)" in s) == consolidated)]
    return (pref or cands)[0]


def _da_from_notes_scan(key, rcept, top, notes_sub):
    """구형 보고서(개별 주석 TITLE 없음): 주석 섹션 전체 표에서 성격별/판관비 표를 찾아 D&A 추출."""
    tables = [p for k, p in dart.get_section_blocks(key, rcept, top, notes_sub)
              if k == "table" and len(p) >= 2]
    nature = sga = None
    for t in tables:
        labels = [(r[0] or "") for r in t]
        j = " ".join(labels)
        if nature is None and (("감가상각비 등" in j) or
                               ("원재료" in j and "종업원급여" in j and any("감가상각" in l for l in labels))):
            nature = _table_da(t)
        if sga is None and ("무형자산상각비" in j) and any("감가상각비" in l for l in labels):
            sga = _table_da(t)
    return nature, sga


def da_breakdown(key, corp_code, year, reprt, fs):
    """당기 D&A(감가상각비+무형자산상각비) 분해. 반환: (전체D&A원, 판관비D&A원) or (None,None).
    전체 = 비용의 성격별 분류, 판관비 = 판매비와관리비 주석. 매출원가D&A = 전체 − 판관비."""
    rcept = dart.find_report_rcept(key, corp_code, year)
    if not rcept:
        return None, None
    toc = dart.get_report_toc(key, rcept)
    fin = next((g for g in toc if "재무에 관한" in g["top"]), None)
    if not fin:
        return None, None
    top, subs, consol = fin["top"], fin["subs"], (fs == "CFS")

    # 1) 개별 주석 TITLE 방식(신형 보고서, 2023~)
    n_nat = _find_note(subs, "비용의 성격별 분류", consol)
    n_sga = _find_note(subs, "판매비와관리비", consol)
    nature = _da_from_first_table(key, rcept, top, n_nat) if n_nat else None
    sga = _da_from_first_table(key, rcept, top, n_sga) if n_sga else None

    # 2) 폴백: 주석 섹션 전체 스캔(구형 보고서, ~2022)
    if nature is None or sga is None:
        notes_sub = _find_note(subs, "재무제표 주석", consol)
        if notes_sub:
            f_nat, f_sga = _da_from_notes_scan(key, rcept, top, notes_sub)
            nature = nature if nature is not None else f_nat
            sga = sga if sga is not None else f_sga

    tot = nature * 1_000_000 if nature is not None else None   # 백만원 → 원
    sg = sga * 1_000_000 if sga is not None else None
    return tot, sg


def _row_label_val(row):
    """행에서 (라벨, 당기값). 라벨=앞쪽 비숫자 셀들 결합(중첩표 대응), 값=첫 숫자셀(당기)."""
    parts, val = [], None
    for c in row:
        n = _num(c)
        if n is None:
            if c and str(c).strip():
                parts.append(str(c).strip())
        else:
            val = n
            break
    return " ".join(parts), val


def _has(label, keywords):
    """공백 무시하고 label에 keywords 중 하나라도 포함되는지('합 계'='합계', '인건비(*1)' 등 대응)."""
    nl = (label or "").replace(" ", "")
    return any(k.replace(" ", "") in nl for k in keywords)


def _table_labels(t):
    return " ".join(_row_label_val(r)[0] for r in t)


def _pick_nature_table(key, rcept, top, sub):
    """성격별 비용 표: 노무 키워드와 감가상각 키워드를 함께 가진 (당기) 첫 표."""
    for k, p in dart.get_section_blocks(key, rcept, top, sub):
        if k == "table" and len(p) >= 3:
            labs = _table_labels(p)
            if _has(labs, _LABOR_KW) and _has(labs, _DA_KW):
                return p
    return None


def _pick_table_with(key, rcept, top, sub, anykw):
    """sub의 표 중 라벨에 anykw 중 하나라도 있는 (당기) 첫 표(>=3행)."""
    for k, p in dart.get_section_blocks(key, rcept, top, sub):
        if k == "table" and len(p) >= 3 and _has(_table_labels(p), anykw):
            return p
    return None


def _rowsum(table, include, exclude=()):
    """표에서 라벨이 include 중 하나 포함(exclude 제외) 행들의 당기값 합계(백만원).
    라벨은 여러 열에서 읽고(중첩표 대응) 값은 첫 숫자(당기)."""
    if not table:
        return None
    tot, found = 0.0, False
    for row in table[1:]:
        lab, v = _row_label_val(row)
        if v is None:
            continue
        if _has(lab, include) and not _has(lab, exclude):
            tot += v
            found = True
    return tot if found else None


# 원가 성격별 라벨 변형(회사·업종마다 다름)
_MAT_KW = ("원재료", "재료비", "원자재", "부재료", "재고자산 매입", "재고자산매입",
           "재고자산의 매입", "상품매입", "매입액")
_LABOR_KW = ("종업원급여", "인건비", "노무비", "급여")
_DA_KW = ("감가상각", "무형자산상각", "무형상각")
_LABOR_SGA_KW = ("급여", "상여", "퇴직", "복리후생", "인건비", "노무")


def cost_breakdown(key, corp_code, year, reprt, fs):
    """성격별 비용·판관비 주석에서 원가 성격별 금액 추출(원).
    반환 dict: total(성격별 합계) material(재료) labor(노무=종업원급여) da_total(감가+무형)
              labor_sga(판관비 인건비) da_sga(판관비 D&A). 못 찾으면 None."""
    rcept = dart.find_report_rcept(key, corp_code, year)
    if not rcept:
        return None
    toc = dart.get_report_toc(key, rcept)
    fin = next((g for g in toc if "재무에 관한" in g["top"]), None)
    if not fin:
        return None
    top, subs, consol = fin["top"], fin["subs"], (fs == "CFS")

    n_nat = (_find_note(subs, "비용의 성격별 분류", consol) or _find_note(subs, "성격별 비용", consol)
             or _find_note(subs, "성격별", consol))
    n_sga = _find_note(subs, "판매비와관리비", consol) or _find_note(subs, "판매비", consol)
    nat_t = _pick_nature_table(key, rcept, top, n_nat) if n_nat else None
    sga_t = _pick_table_with(key, rcept, top, n_sga, _DA_KW) if n_sga else None

    # 폴백: 구형 보고서(개별 주석 TITLE 없음) — 주석 섹션 전체 표 스캔
    if nat_t is None or sga_t is None:
        notes_sub = _find_note(subs, "재무제표 주석", consol)
        if notes_sub:
            tables = [p for k, p in dart.get_section_blocks(key, rcept, top, notes_sub)
                      if k == "table" and len(p) >= 3]
            for t in tables:
                labs = _table_labels(t)
                if nat_t is None and _has(labs, ("합계",)) and _has(labs, _MAT_KW) and _has(labs, _LABOR_KW):
                    nat_t = t
                if sga_t is None and _has(labs, _DA_KW) and _has(labs, ("판매비", "판관비", "급여", "무형자산상각")):
                    sga_t = t

    if nat_t is None:
        return None

    def M(v):
        return v * 1_000_000 if v is not None else None

    return {
        "total": M(_rowsum(nat_t, ("합계",), exclude=("소계",))),
        "material": M(_rowsum(nat_t, _MAT_KW)),
        "labor": M(_rowsum(nat_t, _LABOR_KW)),
        "da_total": M(_rowsum(nat_t, _DA_KW)),
        "labor_sga": M(_rowsum(sga_t, _LABOR_SGA_KW)) if sga_t else None,
        "da_sga": M(_rowsum(sga_t, _DA_KW)) if sga_t else None,
    }


def ppe_depreciation(key, corp_code, year, reprt, fs):
    """유형자산 주석(롤포워드)의 당기 감가상각비(원). 무형상각·감가상각누계액은 제외."""
    rcept = dart.find_report_rcept(key, corp_code, year)
    if not rcept:
        return None
    toc = dart.get_report_toc(key, rcept)
    fin = next((g for g in toc if "재무에 관한" in g["top"]), None)
    if not fin:
        return None
    top, subs, consol = fin["top"], fin["subs"], (fs == "CFS")
    note = _find_note(subs, "유형자산", consol)
    if not note:
        return None
    # 당기 롤포워드 표가 먼저 나오므로 첫 '감가상각비'(누계 제외) 행 = 당기 감가상각비
    for kind, payload in dart.get_section_blocks(key, rcept, top, note):
        if kind != "table":
            continue
        for row in payload[1:]:
            lab = row[0] or ""
            if "감가상각비" in lab and "누계" not in lab:
                v = _first_num(row[1:])
                if v is not None:
                    return v * 1_000_000      # 백만원 → 원
    return None


def _make_row(label, kind, values, years_sorted):
    """연도별 값 + 전년比(금액=증감률, 그외=차이) 계산."""
    changes = {}
    for i in range(1, len(years_sorted)):
        y, py = years_sorted[i], years_sorted[i - 1]
        cur, prev = values.get(y), values.get(py)
        if cur is None or prev is None:
            changes[y] = None
        elif kind == "amount":
            changes[y] = (cur - prev) / abs(prev) if prev != 0 else None
        else:
            changes[y] = cur - prev
    return {"label": label, "kind": kind, "values": values, "changes": changes}


def compute_metrics(key, corp_code, years, reprt, fs, deep=False):
    """회사 1개에 대해 연도별 지표 + 연도마다 전년比 계산.
    deep=True면 주석 기반 '매출원가 중 D&A 비중' 지표까지 추가(문서 파싱, 느림)."""
    years_sorted = sorted(str(y) for y in years)          # 오래된→최근 (왼→오)
    base_by_year = {y: _base_values(dart.get_statement(key, corp_code, y, reprt, fs))
                    for y in years_sorted}

    out_rows = []
    for label, kind, fn in METRIC_DEFS:
        if kind == "sep":
            out_rows.append({"label": "", "kind": "sep", "values": {}, "changes": {}})
            continue
        values = {y: fn(base_by_year[y]) for y in years_sorted}
        out_rows.append(_make_row(label, kind, values, years_sorted))

    if deep:
        ys = years_sorted
        cb = {y: (cost_breakdown(key, corp_code, y, reprt, fs) or {}) for y in ys}

        def F(f):
            return {y: cb[y].get(f) for y in ys}

        def SEP():
            out_rows.append({"label": "", "kind": "sep", "values": {}, "changes": {}})

        def _ratio(num, den):
            return {y: (num[y] / den[y]) if (num[y] is not None and den.get(y)) else None for y in ys}

        def _sub(a, b):
            return {y: (a[y] - b[y]) if (a[y] is not None and b[y] is not None) else None for y in ys}

        total = F("total")          # 성격별 합계(전체 원가)
        material = F("material")     # 재료원가
        labor = F("labor")          # 노무원가(종업원급여)
        da_tot = F("da_total")      # 전체 D&A
        labor_sga = F("labor_sga")  # 판관비 인건비
        da_sga = F("da_sga")        # 판관비 D&A
        cogs_b = {y: base_by_year[y]["cogs"] for y in ys}
        sga_b = {y: base_by_year[y]["sga"] for y in ys}
        # 전체 원가 denominator: 성격별 합계 우선, 없으면 매출원가+판관비
        total_c = {y: total[y] if total[y] is not None else
                   ((cogs_b[y] + sga_b[y]) if (cogs_b[y] is not None and sga_b[y] is not None) else None)
                   for y in ys}

        cogs_da = _sub(da_tot, da_sga)
        mfg_all = {y: (total_c[y] - material[y] - labor[y])
                   if (total_c[y] is not None and material[y] is not None and labor[y] is not None) else None
                   for y in ys}
        cogs_labor = _sub(labor, labor_sga)     # 매출원가 노무 = 전체 노무 − 판관비 노무
        cogs_mfg = {y: (cogs_b[y] - material[y] - cogs_labor[y])
                    if (cogs_b[y] is not None and material[y] is not None and cogs_labor[y] is not None) else None
                    for y in ys}

        # ── 감가상각비·무형상각 D&A 비중
        SEP()
        out_rows.append(_make_row("전체 감가상각비·무형상각(성격별)", "amount", da_tot, ys))
        out_rows.append(_make_row("전체 원가 중 D&A 비중", "pct", _ratio(da_tot, total_c), ys))
        out_rows.append(_make_row("매출원가 감가상각비·무형상각(추정)", "amount", cogs_da, ys))
        out_rows.append(_make_row("매출원가 중 D&A 비중", "pct", _ratio(cogs_da, cogs_b), ys))
        out_rows.append(_make_row("판관비 감가상각비·무형상각", "amount", da_sga, ys))
        out_rows.append(_make_row("판관비 중 D&A 비중", "pct", _ratio(da_sga, sga_b), ys))

        # ── 전체 원가 성격별 구성 (재료·노무·제조간접)
        SEP()
        out_rows.append(_make_row("전체 원가(성격별 합계)", "amount", total_c, ys))
        out_rows.append(_make_row("재료원가", "amount", material, ys))
        out_rows.append(_make_row("전체 원가 중 재료비중", "pct", _ratio(material, total_c), ys))
        out_rows.append(_make_row("노무원가(종업원급여)", "amount", labor, ys))
        out_rows.append(_make_row("전체 원가 중 노무비중", "pct", _ratio(labor, total_c), ys))
        out_rows.append(_make_row("제조간접·기타", "amount", mfg_all, ys))
        out_rows.append(_make_row("전체 원가 중 제조간접비중", "pct", _ratio(mfg_all, total_c), ys))

        # ── 매출원가 성격별 구성
        SEP()
        out_rows.append(_make_row("매출원가 재료원가", "amount", material, ys))
        out_rows.append(_make_row("매출원가 중 재료비중", "pct", _ratio(material, cogs_b), ys))
        out_rows.append(_make_row("매출원가 노무원가", "amount", cogs_labor, ys))
        out_rows.append(_make_row("매출원가 중 노무비중", "pct", _ratio(cogs_labor, cogs_b), ys))
        out_rows.append(_make_row("매출원가 제조간접·기타", "amount", cogs_mfg, ys))
        out_rows.append(_make_row("매출원가 중 제조간접비중", "pct", _ratio(cogs_mfg, cogs_b), ys))

        # ── 재구성 CAPEX = 유형자산 감가상각비(주석) + 유형자산 변동
        SEP()
        ppe_dep = {y: ppe_depreciation(key, corp_code, y, reprt, fs) for y in ys}
        ppe_chg = {y: base_by_year[y]["ppe_chg"] for y in ys}
        recap = {y: (ppe_dep[y] + ppe_chg[y]) if (ppe_dep[y] is not None and ppe_chg[y] is not None) else None
                 for y in ys}
        out_rows.append(_make_row("유형자산 감가상각비(주석)", "amount", ppe_dep, ys))
        out_rows.append(_make_row("유형자산 변동(당기−전기)", "amount", ppe_chg, ys))
        out_rows.append(_make_row("재구성 CAPEX(감가상각비+유형자산변동)", "amount", recap, ys))

    return {"years": years_sorted, "rows": out_rows}

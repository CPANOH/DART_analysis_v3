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


def _find(rows, ids=(), sj=None, name_contains=()):
    """account_id 우선, 없으면 (해당 재무제표 내) 계정명 부분일치로 값 찾기."""
    for r in rows:
        if sj and r.get("sj_div") != sj:
            continue
        if r.get("account_id") in ids:
            return _num(r.get("thstrm_amount"))
    if name_contains:
        for r in rows:
            if sj and r.get("sj_div") != sj:
                continue
            nm = r.get("account_nm") or ""
            if any(k in nm for k in name_contains):
                return _num(r.get("thstrm_amount"))
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
    inv = _find(rows, {"ifrs-full_Inventories"}, "BS", ("재고자산",))
    ocf = _find(rows, {"ifrs-full_CashFlowsFromUsedInOperatingActivities"}, "CF",
                ("영업활동 현금흐름", "영업활동현금흐름", "영업활동으로"))
    capex = _find(rows, {"ifrs-full_PurchaseOfPropertyPlantAndEquipmentClassifiedAsInvestingActivities"},
                  "CF", ("유형자산의 취득", "유형자산 취득", "유형자산의취득"))
    return {"rev": rev, "cogs": cogs, "gp": gp, "sga": sga, "op": op, "ni": ni,
            "ppe": ppe, "inv": inv, "ocf": ocf, "capex": capex}


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
    ("CAPEX(유형자산 취득)", "amount", lambda b: b["capex"]),
    ("단순 FCF(영업CF−CAPEX)", "amount",
     lambda b: (b["ocf"] - b["capex"]) if (b["ocf"] is not None and b["capex"] is not None) else None),
    ("__sep__", "sep", None),
    ("매출원가율", "pct", lambda b: _div(b["cogs"], b["rev"])),
    ("매출총이익률", "pct", lambda b: _div(b["gp"], b["rev"])),
    ("영업이익률", "pct", lambda b: _div(b["op"], b["rev"])),
    ("순이익률", "pct", lambda b: _div(b["ni"], b["rev"])),
    ("CAPEX / 매출액", "pct", lambda b: _div(b["capex"], b["rev"])),
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
        out_rows.append({"label": "", "kind": "sep", "values": {}, "changes": {}})
        tot = {}      # 전체 D&A(원)
        sga = {}      # 판관비 D&A(원)
        for y in years_sorted:
            tot[y], sga[y] = da_breakdown(key, corp_code, y, reprt, fs)
        cogs_da = {y: (tot[y] - sga[y]) if (tot[y] is not None and sga[y] is not None) else None
                   for y in years_sorted}

        def _ratio(num, den):
            return {y: (num[y] / den[y]) if (num[y] is not None and den[y]) else None
                    for y in years_sorted}

        cogs_base = {y: base_by_year[y]["cogs"] for y in years_sorted}
        sga_base = {y: base_by_year[y]["sga"] for y in years_sorted}
        total_cost = {y: (cogs_base[y] + sga_base[y])
                      if (cogs_base[y] is not None and sga_base[y] is not None) else None
                      for y in years_sorted}

        # 매출원가
        out_rows.append(_make_row("매출원가 감가상각비·무형상각(추정)", "amount", cogs_da, years_sorted))
        out_rows.append(_make_row("매출원가 중 D&A 비중", "pct", _ratio(cogs_da, cogs_base), years_sorted))
        # 판관비
        out_rows.append(_make_row("판관비 감가상각비·무형상각", "amount", sga, years_sorted))
        out_rows.append(_make_row("판관비 중 D&A 비중", "pct", _ratio(sga, sga_base), years_sorted))
        # 전체 원가
        out_rows.append(_make_row("전체 감가상각비·무형상각(성격별)", "amount", tot, years_sorted))
        out_rows.append(_make_row("전체 원가 중 D&A 비중", "pct", _ratio(tot, total_cost), years_sorted))

    return {"years": years_sorted, "rows": out_rows}

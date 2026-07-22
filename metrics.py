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


def compute_metrics(key, corp_code, years, reprt, fs):
    """회사 1개에 대해 연도별 지표 + 연도마다 전년比 계산.
    반환: {'years':[...오래된→최근...], 'rows':[{label, kind, values:{year:val}, changes:{year:전년比}}]}
    changes: 금액지표는 증감률(비율), 비율/회전/일수 지표는 전년과의 차이(Δ)."""
    years_sorted = sorted(str(y) for y in years)          # 오래된→최근 (왼→오)
    base_by_year = {}
    for y in years_sorted:
        rows = dart.get_statement(key, corp_code, y, reprt, fs)
        base_by_year[y] = _base_values(rows)

    out_rows = []
    for label, kind, fn in METRIC_DEFS:
        if kind == "sep":
            out_rows.append({"label": "", "kind": "sep", "values": {}, "changes": {}})
            continue
        values = {y: fn(base_by_year[y]) for y in years_sorted}
        changes = {}
        for i in range(1, len(years_sorted)):
            y, py = years_sorted[i], years_sorted[i - 1]
            cur, prev = values[y], values[py]
            if cur is None or prev is None:
                changes[y] = None
            elif kind == "amount":
                changes[y] = (cur - prev) / abs(prev) if prev != 0 else None
            else:                                   # 비율/회전/일수 → 전년과의 차이
                changes[y] = cur - prev
        out_rows.append({"label": label, "kind": kind, "values": values, "changes": changes})
    return {"years": years_sorted, "rows": out_rows}

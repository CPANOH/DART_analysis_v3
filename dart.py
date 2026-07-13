"""DART(전자공시) OpenAPI 래퍼.

- corpCode.xml : 종목코드 -> 고유번호(corp_code) 매핑
- fnlttSinglAcntAll.json : 단일회사 전체 재무제표(계정과목 + 금액)
"""

import io
import json
import zipfile
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

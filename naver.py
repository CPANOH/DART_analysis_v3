"""네이버 증권에서 산업군(업종) 목록과 업종별 종목(회사명 + 종목코드)을 수집한다."""

import re
import requests
from bs4 import BeautifulSoup

BASE = "https://finance.naver.com"
HEADERS = {"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64)"}

# 간단한 메모리 캐시 (프로세스가 살아있는 동안 재요청 방지)
_industries_cache = None
_companies_cache = {}


def _get(url):
    r = requests.get(url, headers=HEADERS, timeout=15)
    r.encoding = "euc-kr"  # 네이버 금융은 EUC-KR 인코딩
    r.raise_for_status()
    return r.text


def get_industries():
    """[{'no': '332', 'name': '종이목재'}, ...] 형태로 업종 목록 반환."""
    global _industries_cache
    if _industries_cache is not None:
        return _industries_cache

    html = _get(f"{BASE}/sise/sise_group.naver?type=upjong")
    soup = BeautifulSoup(html, "html.parser")

    out, seen = [], set()
    for a in soup.select("a[href*=sise_group_detail]"):
        href = a.get("href", "")
        m = re.search(r"no=(\d+)", href)
        name = a.get_text(strip=True)
        if m and name and m.group(1) not in seen:
            seen.add(m.group(1))
            out.append({"no": m.group(1), "name": name})

    out.sort(key=lambda x: x["name"])
    _industries_cache = out
    return out


def get_companies(no):
    """특정 업종(no)에 속한 종목 목록. [{'name': '모나미', 'stock_code': '005360'}, ...]"""
    if no in _companies_cache:
        return _companies_cache[no]

    html = _get(f"{BASE}/sise/sise_group_detail.naver?type=upjong&no={no}")
    soup = BeautifulSoup(html, "html.parser")

    out, seen = [], set()
    for a in soup.select('a[href*="item/main"]'):
        href = a.get("href", "")
        m = re.search(r"code=(\d{6})", href)
        name = a.get_text(strip=True)
        if m and name and m.group(1) not in seen:
            seen.add(m.group(1))
            out.append({"name": name, "stock_code": m.group(1)})

    out.sort(key=lambda x: x["name"])
    _companies_cache[no] = out
    return out

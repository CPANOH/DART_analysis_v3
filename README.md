# DART_analysis_v3 · 재무분석기

> This tool's purpose is to analyze F/S from DART for studying business, value chain, etc.

산업군 → 회사명 → 계정과목을 순서대로 필터링하고, 선택한 재무 데이터를 **엑셀(.xlsx)** 로 내려받는 웹앱입니다.

- **1차 필터 · 산업군** — [네이버 증권](https://finance.naver.com) 업종 분류
- **2차 필터 · 회사명** — 네이버 업종별 종목 + [DART](https://opendart.fss.or.kr) 고유번호 매칭
- **3차 필터 · 계정과목** — DART 재무제표 API(`fnlttSinglAcntAll`), 재무제표 표시 순서로 정렬
- **최대 5개 연도** 다중 선택 → 회사×연도 비교표
- **플러스알파**: 사업보고서 주요정보(배당·임원·주주 등)와 본문 서술 섹션(대분류→소분류)도 엑셀 시트로

## 필요한 것

- Python 3.10+
- **DART OpenAPI 인증키** (무료) — https://opendart.fss.or.kr 에서 발급

## 로컬 실행

```bash
pip install -r requirements.txt
python app.py
```

브라우저가 자동으로 `http://127.0.0.1:5000` 을 엽니다.
DART 키는 화면 상단에 직접 입력하거나, 로컬 `.env` 파일 / 환경변수 `DART_API_KEY` 로 지정할 수 있습니다.

Windows에서는 `run.bat` 을 더블클릭해도 됩니다.

## Vercel 배포

1. 이 저장소를 GitHub에 올립니다.
2. [Vercel](https://vercel.com) 에서 **Add New → Project → Import** 로 저장소를 선택합니다.
3. **Settings → Environment Variables** 에서 다음을 추가합니다.
   - `DART_API_KEY` = 발급받은 DART 키
4. **Deploy** 를 누르면 완료됩니다.

> ⚠️ **보안**: DART 키는 절대 코드에 하드코딩하지 마세요. 반드시 환경변수로만 사용합니다.
> 서버에 `DART_API_KEY` 를 설정하면 방문자는 키 없이 사용할 수 있으나, 그만큼 본인 키의
> 호출 한도(일 20,000회)를 공유하게 됩니다. 원치 않으면 환경변수를 비워두고 각자
> 화면에서 키를 입력하도록 하세요.

## 구조

```
app.py                Flask 서버 + 엑셀 생성
naver.py              네이버 증권 업종/종목 수집
dart.py               DART API 래퍼 (고유번호, 재무제표, 사업보고서 본문)
templates/index.html  프론트엔드(필터 UI)
api/index.py          Vercel 서버리스 진입점
vercel.json           Vercel 설정
main.py               (참고용) CLI 버전 재무 비교 스크립트
```

데이터 출처: 네이버 증권, 금융감독원 전자공시시스템(DART).

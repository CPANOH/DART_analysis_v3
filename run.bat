@echo off
chcp 65001 >nul
cd /d "%~dp0"
echo ============================================
echo   재무분석기 실행
echo ============================================
echo.
echo [1/2] 필요한 패키지 확인/설치 중...
python -m pip install -q -r requirements.txt
echo [2/2] 서버 시작... (브라우저가 자동으로 열립니다)
echo.
python app.py
pause

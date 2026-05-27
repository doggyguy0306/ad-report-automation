@echo off
REM ─────────────────────────────────────────────
REM  광고 성과 자동화 — 서버 시작 스크립트 (Windows)
REM  보고서를 만들 날에 이 파일을 더블클릭하세요.
REM  창을 닫으면 서버가 종료됩니다.
REM ─────────────────────────────────────────────
cd /d "%~dp0"

echo.
echo ================================================
echo   광고 성과 자동화  ^|  서버 시작
echo ================================================
echo.

REM IP 주소 표시
for /f "tokens=2 delims=:" %%a in ('ipconfig ^| findstr "IPv4"') do (
    set RAW=%%a
    goto :found
)
:found
set IP=%RAW: =%

echo   팀원 접속 주소:
echo.
echo        http://%IP%:8501
echo.
echo   이 창을 닫으면 서버가 종료됩니다.
echo ================================================
echo.

REM Streamlit 서버 시작
python -m streamlit run app.py ^
    --server.address 0.0.0.0 ^
    --server.port 8501 ^
    --server.headless true ^
    --browser.gatherUsageStats false

pause

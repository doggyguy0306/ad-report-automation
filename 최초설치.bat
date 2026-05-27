@echo off
REM ─────────────────────────────────────────────
REM  광고 성과 자동화 — 최초 설치 스크립트 (Windows)
REM  이 파일을 더블클릭해서 한 번만 실행하세요.
REM ─────────────────────────────────────────────
cd /d "%~dp0"

echo.
echo ================================================
echo   광고 성과 자동화  ^|  최초 설치
echo ================================================
echo.

REM Python 확인
python --version >nul 2>&1
if errorlevel 1 (
    echo ❌ Python이 설치되어 있지 않습니다.
    echo    https://www.python.org/downloads/ 에서 설치 후
    echo    다시 실행하세요.
    echo.
    pause
    exit /b 1
)

echo 필요한 패키지를 설치합니다...
echo.
pip install -r requirements.txt

echo.
echo ================================================
echo   설치 완료!
echo.
echo   이제 '서버_시작.bat' 를 더블클릭하면
echo   서버가 시작됩니다.
echo ================================================
echo.
pause

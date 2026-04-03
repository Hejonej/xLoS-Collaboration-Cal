@echo off
chcp 65001 >nul
echo.
echo   ╔══════════════════════════════════════╗
echo   ║   xLoS / 동일LoS 분석 도구          ║
echo   ╚══════════════════════════════════════╝
echo.

cd /d "%~dp0"

pip install -q -r requirements.txt 2>nul

echo   [1] 내 PC에서만 사용  (localhost)
echo   [2] 사내 네트워크 공유 (동료 접속 가능)
echo.
set /p MODE="   선택 (1 또는 2): "

if "%MODE%"=="2" (
    echo.
    echo   ── 사내 네트워크 모드 ──
    echo.
    for /f "tokens=2 delims=:" %%a in ('ipconfig ^| findstr /i "IPv4"') do (
        for /f "tokens=1" %%b in ("%%a") do (
            echo   동료 접속 주소:  http://%%b:5000
        )
    )
    echo   내 PC 접속:      http://localhost:5000
    echo.
    echo   종료: Ctrl+C
    echo   ════════════════════════════════════════
    python -c "from waitress import serve; from app import app; print(); serve(app, host='0.0.0.0', port=5000, threads=4)"
) else (
    echo.
    echo   ── 로컬 모드 ──
    echo   접속 주소: http://localhost:5000
    echo.
    echo   종료: Ctrl+C
    echo   ════════════════════════════════════════
    python app.py
)

pause

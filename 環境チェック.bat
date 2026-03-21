@echo off
echo.
echo =============================================
echo   環境チェック
echo =============================================
echo.
cd /d "%~dp0"
python generate.py --check
echo.
pause

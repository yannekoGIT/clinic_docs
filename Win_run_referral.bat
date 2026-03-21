@echo off
echo.
echo Referral batch generation
echo.
cd /d "%~dp0"
if "%1"=="--check" (
    python generate.py --check
    pause
    exit /b
)
python run_batch.py referral --save-json %*
echo.
pause

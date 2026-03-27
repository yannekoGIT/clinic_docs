@echo off
echo.
echo =============================================
echo   障害年金診断書（精神の障害用）バッチ作成
echo =============================================
echo.
cd /d "%~dp0"
if "%1"=="--check" (
    python generate.py --check
    echo.
    pause
    exit /b
)
where python >nul 2>&1
if %ERRORLEVEL% NEQ 0 (
    echo [ERROR] Python が見つかりません。
    pause
    exit /b 1
)
echo 障害年金診断書（.xlsx）を生成します。
echo.
python run_batch.py shougai_nenkin %*
echo.
if %ERRORLEVEL% EQU 0 (
    echo 出力先: output フォルダを確認してください。
) else (
    echo エラーが発生しました。
)
echo.
pause

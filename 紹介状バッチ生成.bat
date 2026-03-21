@echo off
echo.
echo =============================================
echo   紹介状（診療情報提供書）バッチ生成
echo =============================================
echo.
cd /d "%~dp0"
if "%1"=="--check" (
    python generate.py --check
    echo.
    pause
    exit /b
)
where python >/dev/null 2>&1
if %ERRORLEVEL% NEQ 0 (
    echo [ERROR] Python が見つかりません。
    pause
    exit /b 1
)
where node >/dev/null 2>&1
if %ERRORLEVEL% NEQ 0 (
    echo [ERROR] Node.js が見つかりません。
    pause
    exit /b 1
)
echo 紹介状（.docx）を生成します。
echo.
python run_batch.py referral --save-json %*
echo.
if %ERRORLEVEL% EQU 0 (
    echo 出力先: output フォルダを確認してください。
) else (
    echo エラーが発生しました。
)
echo.
pause

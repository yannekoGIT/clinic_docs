@echo off
echo.
echo =============================================
echo   紹介状（診療情報提供書）作成
echo   コピペ用テキスト（.txt）を生成します
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
python run_batch.py referral %*
echo.
if %ERRORLEVEL% EQU 0 (
    echo 出力先: output フォルダを確認してください。
) else (
    echo エラーが発生しました。
)
echo.
pause

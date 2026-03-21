@echo off
echo.
echo =============================================
echo   自立支援医療診断書（精神通院医療用）作成
echo   Word版（.docx）で生成します
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
python run_batch.py jiritsu_shien_word --save-json %*
echo.
if %ERRORLEVEL% EQU 0 (
    echo 出力先: output フォルダを確認してください。
) else (
    echo エラーが発生しました。
)
echo.
pause

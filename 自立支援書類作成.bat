@echo off
echo.
echo =============================================
echo   自立支援医療診断書（精神通院医療用）作成
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
echo Word版（.docx）で生成します。
echo Excel版で生成する場合は 2 を入力してください。
echo.
set /p FORMAT_CHOICE=Enter で続行 / 2 でExcel版: 
if "%FORMAT_CHOICE%"=="2" (
    echo.
    echo Excel版で生成します。
    echo.
    python run_batch.py jiritsu_shien --save-json %*
) else (
    echo.
    echo Word版で生成します。
    echo.
    python run_batch.py jiritsu_shien_word --save-json %*
)
echo.
if %ERRORLEVEL% EQU 0 (
    echo 出力先: output フォルダを確認してください。
) else (
    echo エラーが発生しました。
)
echo.
pause

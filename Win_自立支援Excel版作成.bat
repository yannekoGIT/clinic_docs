@echo off
echo.
echo =============================================
echo   自立支援医療診断書（精神通院医療用）作成
echo   Excel版（.xls）で生成します
echo =============================================
echo.
cd /d "%~dp0"
if "%1"=="--check" goto :do_check
where python >nul 2>&1
if %ERRORLEVEL% NEQ 0 goto :no_python
echo 自立支援医療診断書（.xls）を生成します。
echo.
python run_batch.py jiritsu_shien %*
echo.
if %ERRORLEVEL% NEQ 0 goto :error
echo 出力先: output フォルダを確認してください。
echo.
pause
exit /b 0

:do_check
python generate.py --check
echo.
pause
exit /b

:no_python
echo [ERROR] Python が見つかりません。
pause
exit /b 1

:error
echo エラーが発生しました。
echo.
pause
exit /b 1

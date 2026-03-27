@echo off
cd /d "%~dp0"
echo.
echo =============================================
echo   セットアップ / アップデート
echo =============================================
echo.

if exist config.json goto :config_exists
if not exist config.example.json goto :no_example
copy config.example.json config.json >nul
echo [初回] config.json を作成しました。
echo        クリニック情報・API設定を編集してください。
echo.
goto :check_python

:no_example
echo [ERROR] config.example.json が見つかりません。
pause
exit /b 1

:config_exists
echo [OK] config.json は既存のものを保持します。

:check_python
where python >/dev/null 2>&1
if %ERRORLEVEL% NEQ 0 goto :no_python
for /f "tokens=*" %%i in ('python --version 2^>^&1') do echo [OK] %%i
goto :check_node

:no_python
echo [NG] Python が見つかりません。
echo     https://www.python.org/downloads/ からインストールしてください。
pause
exit /b 1

:check_node
where node >/dev/null 2>&1
if %ERRORLEVEL% NEQ 0 goto :no_node
for /f "tokens=*" %%i in ('node --version 2^>^&1') do echo [OK] Node.js %%i
goto :install_deps

:no_node
echo [NG] Node.js が見つかりません。
echo     https://nodejs.org/ からインストールしてください。
pause
exit /b 1

:install_deps
echo.
echo --- Python パッケージをインストール中... ---
pip install -r requirements.txt --quiet
if %ERRORLEVEL% NEQ 0 goto :pip_fail
echo [OK] Python パッケージ完了
echo.
echo --- Node.js パッケージをインストール中... ---
call npm install --silent 2>nul
if %ERRORLEVEL% NEQ 0 goto :npm_fail
echo [OK] Node.js パッケージ完了

if not exist output mkdir output

echo.
echo =============================================
echo   セットアップ完了!
echo =============================================
echo.
echo   初回の場合: config.json を編集してください。
echo   アップデートの場合: 設定はそのまま保持されています。
echo.
pause
exit /b 0

:pip_fail
echo [NG] pip install に失敗しました。
pause
exit /b 1

:npm_fail
echo [NG] npm install に失敗しました。
pause
exit /b 1

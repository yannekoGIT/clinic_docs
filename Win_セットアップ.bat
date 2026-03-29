@echo off
cd /d "%~dp0"
echo.
echo =============================================
echo   セットアップ / アップデート
echo =============================================
echo.

REM ==========================================
REM  1. config.json
REM ==========================================
if exist config.json goto :config_exists
if not exist config.example.json goto :no_example
copy config.example.json config.json >NUL
echo [初回] config.json を作成しました。
echo        クリニック情報を編集してください。
echo.
goto :check_python

:no_example
echo [ERROR] config.example.json が見つかりません。
pause
exit /b 1

:config_exists
echo [OK] config.json は既存のものを保持します。

REM ==========================================
REM  2. Python チェック
REM ==========================================
:check_python
where python >NUL 2>&1
if %ERRORLEVEL% NEQ 0 goto :no_python
for /f "tokens=*" %%i in ('python --version 2^>^&1') do echo [OK] %%i
goto :check_node

:no_python
echo.
echo  ***********************************************
echo  *  Python が見つかりません                     *
echo  ***********************************************
echo.
echo   1. ブラウザで以下を開く:
echo      https://www.python.org/downloads/
echo.
echo   2.「Download Python」ボタンをクリック
echo   3. ダウンロードしたファイルを実行
echo   4.「Add python.exe to PATH」に必ずチェックを入れる
echo   5.「Install Now」をクリック
echo   6. インストール完了後、このバッチを再実行してください
echo.
pause
exit /b 1

REM ==========================================
REM  3. Node.js チェック
REM ==========================================
:check_node
where node >NUL 2>&1
if %ERRORLEVEL% NEQ 0 goto :no_node
for /f "tokens=*" %%i in ('node --version 2^>^&1') do echo [OK] Node.js %%i
goto :check_codex

:no_node
echo.
echo  ***********************************************
echo  *  Node.js が見つかりません                    *
echo  ***********************************************
echo.
echo   1. ブラウザで以下を開く:
echo      https://nodejs.org/
echo.
echo   2.「LTS」版のダウンロードボタンをクリック
echo   3. ダウンロードしたファイルを実行
echo   4. 設定はすべてデフォルト（そのまま Next）でOK
echo   5. インストール完了後、このバッチを再実行してください
echo.
pause
exit /b 1

REM ==========================================
REM  4. Codex CLI チェック
REM ==========================================
:check_codex
where codex >NUL 2>&1
if %ERRORLEVEL% NEQ 0 goto :install_codex
for /f "tokens=*" %%i in ('codex --version 2^>^&1') do echo [OK] Codex CLI %%i
goto :check_codex_auth

:install_codex
echo.
echo  --- Codex CLI をインストール中... ---
call npm install -g @openai/codex 2>NUL
if %ERRORLEVEL% NEQ 0 goto :codex_install_fail
echo  [OK] Codex CLI インストール完了
goto :check_codex_auth

:codex_install_fail
echo.
echo  ***********************************************
echo  *  Codex CLI のインストールに失敗しました      *
echo  ***********************************************
echo.
echo   以下を手動で試してください:
echo     npm install -g @openai/codex
echo.
echo   ※ Codex CLI が無くても config.json の provider を
echo      "api" に変更すれば OpenAI API 経由で利用できます。
echo.
pause
exit /b 1

REM ==========================================
REM  5. Codex CLI 認証チェック
REM ==========================================
:check_codex_auth
REM OPENAI_API_KEY が設定済みならOK
if not "%OPENAI_API_KEY%"=="" goto :codex_auth_ok

REM ~/.codex/ に設定ファイルがあればOK
if exist "%USERPROFILE%\.codex\config.toml" goto :codex_auth_ok

echo.
echo  ***********************************************
echo  *  Codex CLI の認証設定が必要です              *
echo  ***********************************************
echo.
echo   OpenAI の API キーを設定します。
echo.
echo   【APIキーの取得方法】
echo     1. ブラウザで以下を開く:
echo        https://platform.openai.com/api-keys
echo.
echo     2. OpenAI アカウントでログイン
echo        （アカウントが無い場合は「Sign up」から作成）
echo.
echo     3.「Create new secret key」をクリック
echo     4. 表示されたキー（sk-... で始まる文字列）をコピー
echo.
echo   ※ APIの利用には支払い設定が必要です（従量課金）
echo     https://platform.openai.com/settings/organization/billing
echo.
echo  ---------------------------------------------
echo.

set /p "USER_API_KEY=  APIキーを貼り付けて Enter: "

if "%USER_API_KEY%"=="" goto :codex_auth_skip

REM システム環境変数に永続設定
setx OPENAI_API_KEY "%USER_API_KEY%" >NUL 2>&1
if %ERRORLEVEL% NEQ 0 goto :setx_fail

REM 現在のセッションにも反映
set "OPENAI_API_KEY=%USER_API_KEY%"

echo.
echo  [OK] APIキーを設定しました（次回以降も有効です）
goto :codex_auth_ok

:setx_fail
echo.
echo  [注意] 環境変数の永続設定に失敗しました。
echo         手動で設定してください:
echo           setx OPENAI_API_KEY "sk-..."
echo.
echo         今回のセッションでは一時的に使用します。
set "OPENAI_API_KEY=%USER_API_KEY%"
goto :codex_auth_ok

:codex_auth_skip
echo.
echo  [スキップ] APIキーは後から設定できます。
echo.
echo   設定方法:
echo     1. コマンドプロンプトで: setx OPENAI_API_KEY "sk-..."
echo     2. または config.json の provider を "api" に変更し、
echo        環境変数 OPENAI_API_KEY を設定
echo.

:codex_auth_ok
echo.

REM ==========================================
REM  6. 依存パッケージ
REM ==========================================
echo --- Python パッケージをインストール中... ---
python -m pip install -r requirements.txt --quiet
if %ERRORLEVEL% NEQ 0 goto :pip_fail
echo [OK] Python パッケージ完了
echo.
echo --- Node.js パッケージをインストール中... ---
call npm install --silent 2>NUL
if %ERRORLEVEL% NEQ 0 goto :npm_fail
echo [OK] Node.js パッケージ完了

if not exist output mkdir output

echo.
echo =============================================
echo   セットアップ完了!
echo =============================================
echo.
echo   初回の場合:
echo     1. config.json のクリニック情報を編集してください
echo     2. input フォルダにカルテ(.txt)を配置して書類生成バッチを実行
echo.
echo   アップデートの場合:
echo     設定はそのまま保持されています。
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

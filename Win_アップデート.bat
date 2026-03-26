@echo off
chcp 65001 >nul
cd /d "%~dp0"
echo.
echo =============================================
echo   オンラインアップデート
echo =============================================
echo.

:: ---- 1. git チェック ----
where git >nul 2>&1
if %ERRORLEVEL% NEQ 0 (
    echo [ERROR] git が見つかりません。
    echo         https://git-scm.com/downloads からインストールしてください。
    echo.
    pause
    exit /b 1
)

:: ---- 2. gitリポジトリかチェック ----
git rev-parse --git-dir >nul 2>&1
if %ERRORLEVEL% NEQ 0 (
    echo [ERROR] このフォルダは git リポジトリではありません。
    echo         初回セットアップがまだの場合は、以下を実行してください:
    echo.
    echo         git clone https://github.com/yannekoGIT/clinic_docs.git
    echo.
    pause
    exit /b 1
)

:: ---- 3. git pull ----
echo --- 最新版を取得中... ---
git pull origin master
if %ERRORLEVEL% NEQ 0 (
    echo.
    echo [ERROR] 取得に失敗しました。
    echo         ネットワーク接続またはGitHub認証を確認してください。
    echo.
    echo   認証の設定方法:
    echo     git config --global credential.helper manager
    echo     （次回の git pull 時にブラウザでログインを求められます）
    echo.
    pause
    exit /b 1
)

:: ---- 4. 依存パッケージ更新 ----
echo.
echo --- Python パッケージを更新中... ---
pip install -r requirements.txt --quiet
echo [OK] Python パッケージ完了

echo.
echo --- Node.js パッケージを更新中... ---
call npm install --silent 2>nul
echo [OK] Node.js パッケージ完了

:: ---- 5. 完了 ----
echo.
echo =============================================
echo   アップデート完了！
echo =============================================
echo.
echo   config.json の設定はそのまま保持されています。
echo.
pause

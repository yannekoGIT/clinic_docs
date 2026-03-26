@echo off
chcp 65001 >nul
cd /d "%~dp0"
echo.
echo =============================================
echo   セットアップ / アップデート
echo =============================================
echo.

:: ---- 1. config.json が無ければテンプレからコピー ----
if not exist config.json (
    if exist config.example.json (
        copy config.example.json config.json >nul
        echo [初回] config.json を作成しました。
        echo        クリニック情報・API設定を編集してください。
        echo.
    ) else (
        echo [ERROR] config.example.json が見つかりません。
        pause
        exit /b 1
    )
) else (
    echo [OK] config.json は既存のものを保持します。
)

:: ---- 2. Python チェック ----
where python >nul 2>&1
if %ERRORLEVEL% NEQ 0 (
    echo [NG] Python が見つかりません。
    echo     https://www.python.org/downloads/ からインストールしてください。
    pause
    exit /b 1
)
for /f "tokens=*" %%i in ('python --version 2^>^&1') do echo [OK] %%i

:: ---- 3. Node.js チェック ----
where node >nul 2>&1
if %ERRORLEVEL% NEQ 0 (
    echo [NG] Node.js が見つかりません。
    echo     https://nodejs.org/ からインストールしてください。
    pause
    exit /b 1
)
for /f "tokens=*" %%i in ('node --version 2^>^&1') do echo [OK] Node.js %%i

:: ---- 4. 依存パッケージインストール ----
echo.
echo --- Python パッケージをインストール中... ---
pip install -r requirements.txt --quiet
if %ERRORLEVEL% NEQ 0 (
    echo [NG] pip install に失敗しました。
    pause
    exit /b 1
)
echo [OK] Python パッケージ完了

echo.
echo --- Node.js パッケージをインストール中... ---
call npm install --silent 2>nul
if %ERRORLEVEL% NEQ 0 (
    echo [NG] npm install に失敗しました。
    pause
    exit /b 1
)
echo [OK] Node.js パッケージ完了

:: ---- 5. 出力フォルダ作成 ----
if not exist output mkdir output

:: ---- 6. 完了 ----
echo.
echo =============================================
echo   セットアップ完了！
echo =============================================
echo.
echo   初回の場合: config.json を編集してください。
echo   アップデートの場合: 設定はそのまま保持されています。
echo.
pause

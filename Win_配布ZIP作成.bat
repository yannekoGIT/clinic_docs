@echo off
chcp 65001 >nul
cd /d "%~dp0"
echo.
echo =============================================
echo   配布用 ZIP 作成
echo =============================================
echo.

:: ---- git archive で追跡ファイルのみZIP化 ----
where git >nul 2>&1
if %ERRORLEVEL% NEQ 0 (
    echo [ERROR] git が見つかりません。Git をインストールしてください。
    pause
    exit /b 1
)

:: 出力先: 一つ上のフォルダに clinic_docs.zip
set "ZIPFILE=%~dp0..\clinic_docs.zip"

git archive --format=zip --output="%ZIPFILE%" HEAD
if %ERRORLEVEL% NEQ 0 (
    echo [ERROR] ZIP作成に失敗しました。
    pause
    exit /b 1
)

echo [OK] 作成完了: %ZIPFILE%
echo.
echo --- 含まれないもの（自動除外） ---
echo   config.json     ... 各PCの設定を保持するため
echo   node_modules/   ... セットアップ時に自動インストール
echo   output/         ... 生成済みファイル
echo   __pycache__/    ... Pythonキャッシュ
echo.
echo --- 別のPCでの使い方 ---
echo   1. ZIPを展開（既存フォルダに上書き展開OK）
echo   2. Win_セットアップ.bat をダブルクリック
echo      → config.json が無ければ自動作成、依存パッケージも自動インストール
echo      → config.json が既にあれば設定はそのまま保持
echo.
pause

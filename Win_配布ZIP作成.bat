@echo off
cd /d "%~dp0"
echo.
echo =============================================
echo   配布用 ZIP 作成
echo =============================================
echo.

where git >/dev/null 2>&1
if %ERRORLEVEL% NEQ 0 goto :no_git

set "ZIPFILE=%~dp0..\clinic_docs.zip"

git archive --format=zip --output="%ZIPFILE%" HEAD
if %ERRORLEVEL% NEQ 0 goto :zip_fail

echo [OK] 作成完了: %ZIPFILE%
echo.
echo --- 含まれないもの ---
echo   config.json     ... 各PCの設定を保持するため
echo   node_modules/   ... セットアップ時に自動インストール
echo   output/         ... 生成済みファイル
echo   __pycache__/    ... Pythonキャッシュ
echo.
echo --- 別のPCでの使い方 ---
echo   1. ZIPを展開して上書きOK
echo   2. Win_セットアップ.bat をダブルクリック
echo.
pause
exit /b 0

:no_git
echo [ERROR] git が見つかりません。Git をインストールしてください。
pause
exit /b 1

:zip_fail
echo [ERROR] ZIP作成に失敗しました。
pause
exit /b 1

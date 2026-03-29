@echo off
cd /d "%~dp0"
echo.
echo =============================================
echo   オンラインアップデート
echo =============================================
echo.

where git >NUL 2>&1
if %ERRORLEVEL% NEQ 0 goto :no_git

git rev-parse --git-dir >NUL 2>&1
if %ERRORLEVEL% NEQ 0 goto :no_repo

REM --- ローカル変更チェック ---
git diff --quiet >NUL 2>&1
if %ERRORLEVEL% NEQ 0 goto :local_changes
git diff --cached --quiet >NUL 2>&1
if %ERRORLEVEL% NEQ 0 goto :local_changes

echo --- 最新版を取得中... ---
git pull origin master
if %ERRORLEVEL% NEQ 0 goto :pull_fail

goto :install_deps

REM ==========================================
REM  git pull 失敗 : 原因切り分け
REM ==========================================
:pull_fail
echo.

REM ネットワーク疎通チェック
ping -n 1 -w 3000 github.com >NUL 2>&1
if %ERRORLEVEL% NEQ 0 goto :network_error

REM ネットワークOK : 認証の問題
goto :auth_error

REM ==========================================
REM  ネットワークエラー
REM ==========================================
:network_error
echo.
echo  ***********************************************
echo  *  ネットワークに接続できません               *
echo  ***********************************************
echo.
echo   以下を確認してください:
echo.
echo   1. Wi-Fi またはLANケーブルが接続されていますか？
echo   2. ブラウザで適当なサイトは開けますか？
echo   3. 病院のファイアウォールでブロックされていませんか？
echo      → 院内ネットワーク担当者にご相談ください
echo.
pause
exit /b 1

REM ==========================================
REM  認証エラー : 自動設定＋リトライ
REM ==========================================
:auth_error
echo.
echo  ***********************************************
echo  *  GitHubへのログインが必要です                *
echo  ***********************************************
echo.
echo   このパソコンではまだGitHubの認証が
echo   設定されていないようです。
echo   以下の手順で設定します（初回のみ）。
echo.

REM credential helper が未設定なら自動設定
set "CRED_HELPER="
for /f "tokens=*" %%a in ('git config --global credential.helper 2^>NUL') do set "CRED_HELPER=%%a"
if "%CRED_HELPER%"=="" goto :setup_credential
goto :check_github_account

:setup_credential
echo  [自動設定] 認証ヘルパーを設定しています...
git config --global credential.helper manager
echo  [OK] 設定完了（Windows資格情報マネージャーを使用）
echo.

:check_github_account
echo  ---------------------------------------------
echo    GitHubアカウントをお持ちですか？
echo  ---------------------------------------------
echo.
echo   【持っている場合】
echo     → このまま Enter を押してください。
echo       ブラウザが開くので、GitHubにログインしてください。
echo.
echo   【持っていない場合】
echo     → 先にアカウントを作成してください:
echo.
echo     1. ブラウザで以下を開く:
echo        https://github.com/signup
echo.
echo     2. メールアドレスを入力
echo     3. パスワードを設定（15文字以上、または8文字以上で
echo        数字と小文字を含む）
echo     4. ユーザー名を決める（半角英数字）
echo     5. メールに届く認証コードを入力
echo.
echo     ※ 無料プランで問題ありません
echo     ※ アカウント作成後、ここに戻って Enter を押してください
echo.
echo  ---------------------------------------------
echo.
pause

echo.
echo --- 再度取得を試みています... ---
echo    （ブラウザが開いたらGitHubにログインしてください）
echo.
git pull origin master
if %ERRORLEVEL% NEQ 0 goto :auth_retry_fail

goto :install_deps

REM ==========================================
REM  リトライも失敗
REM ==========================================
:auth_retry_fail
echo.
echo  ***********************************************
echo  *  再試行でも取得できませんでした              *
echo  ***********************************************
echo.
echo   考えられる原因:
echo.
echo   1. GitHubのログイン画面でキャンセルした
echo      → もう一度このバッチファイルを実行してください
echo.
echo   2. アカウントはあるがリポジトリへのアクセス権がない
echo      → 管理者にGitHubユーザー名を伝えて、
echo        招待してもらってください
echo.
echo   3. ブラウザでのログインがうまくいかない
echo      → 以下を試してください:
echo         a. ブラウザで https://github.com にログインできるか確認
echo         b. ログインできたら、このバッチを再実行
echo.
echo   それでも解決しない場合はシステム管理者にご連絡ください。
echo.
pause
exit /b 1

REM ==========================================
REM  ローカル変更あり
REM ==========================================
:local_changes
echo  ***********************************************
echo  *  ローカルに変更されたファイルがあります       *
echo  ***********************************************
echo.
echo   アップデートするとローカルの変更が競合する場合が
echo   あります。変更を一時退避してからアップデートします。
echo.

git stash
if %ERRORLEVEL% NEQ 0 goto :stash_fail

echo  [OK] ローカル変更を一時退避しました
echo.
echo --- 最新版を取得中... ---
git pull origin master
if %ERRORLEVEL% NEQ 0 goto :pull_fail_after_stash

echo.
echo --- 退避した変更を復元中... ---
git stash pop >NUL 2>&1
if %ERRORLEVEL% NEQ 0 goto :stash_conflict
echo  [OK] ローカル変更を復元しました

goto :install_deps

:stash_conflict
echo  [注意] 一部のファイルが競合しました。
echo         手動での確認が必要です。
echo         退避データは git stash list で確認できます。
goto :install_deps

:stash_fail
echo  [ERROR] ローカル変更の退避に失敗しました。
echo          手動で確認してください。
pause
exit /b 1

:pull_fail_after_stash
echo.
echo  [注意] 取得に失敗しましたが、退避した変更は保持されています。
echo         git stash pop で復元できます。
echo.

REM stash後のpull失敗も同じ切り分けを行う
ping -n 1 -w 3000 github.com >NUL 2>&1
if %ERRORLEVEL% NEQ 0 goto :network_error
goto :auth_error

REM ==========================================
REM  依存パッケージ更新
REM ==========================================
:install_deps
echo.
echo --- Python パッケージを更新中... ---
python -m pip install -r requirements.txt --quiet
echo [OK] Python パッケージ完了
echo.
echo --- Node.js パッケージを更新中... ---
call npm install --silent 2>NUL
echo [OK] Node.js パッケージ完了

echo.
echo =============================================
echo   アップデート完了!
echo =============================================
echo.
echo   config.json の設定はそのまま保持されています。
echo.
pause
exit /b 0

REM ==========================================
REM  git未インストール
REM ==========================================
:no_git
echo  ***********************************************
echo  *  git が見つかりません                        *
echo  ***********************************************
echo.
echo   Gitのインストールが必要です:
echo.
echo   1. ブラウザで以下を開く:
echo      https://git-scm.com/downloads/win
echo.
echo   2.「Click here to download」をクリック
echo   3. ダウンロードしたファイルを実行
echo   4. 設定はすべてデフォルト（そのまま Next）でOK
echo   5. インストール完了後、このバッチを再実行してください
echo.
pause
exit /b 1

REM ==========================================
REM  gitリポジトリではない
REM ==========================================
:no_repo
echo  ***********************************************
echo  *  このフォルダはgitリポジトリではありません     *
echo  ***********************************************
echo.
echo   初回セットアップが必要です。
echo   セットアップ用のバッチファイルを実行するか、
echo   以下のコマンドを実行してください:
echo.
echo   git clone https://github.com/yannekoGIT/clinic_docs.git
echo.
pause
exit /b 1

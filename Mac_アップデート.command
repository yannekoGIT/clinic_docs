#!/bin/bash
# オンラインアップデート
cd "$(dirname "$0")"
echo ""
echo "============================================="
echo "  オンラインアップデート"
echo "============================================="
echo ""

# ---- 1. git チェック ----
if ! command -v git &>/dev/null; then
    echo "[ERROR] git が見つかりません。"
    echo "        xcode-select --install または https://git-scm.com/downloads"
    read -p "Press Enter to close..."
    exit 1
fi

# ---- 2. gitリポジトリかチェック（未初期化なら自動設定） ----
if ! git rev-parse --git-dir &>/dev/null; then
    echo "  --- gitリポジトリを初期化しています... ---"
    echo ""
    git init >/dev/null 2>&1
    if [ $? -ne 0 ]; then
        echo "[ERROR] gitリポジトリの初期化に失敗しました。"
        read -p "Press Enter to close..."
        exit 1
    fi
    git remote add origin https://github.com/yannekoGIT/clinic_docs.git 2>/dev/null
    echo "  [OK] git初期化完了"
    echo ""
    echo "--- 最新版を取得中... ---"
    git fetch origin master >/dev/null 2>&1
    git reset origin/master >/dev/null 2>&1
fi

# ---- 3. git pull ----
echo "--- 最新版を取得中... ---"
git pull origin master
if [ $? -ne 0 ]; then
    echo ""
    echo "[ERROR] 取得に失敗しました。"
    echo "        ネットワーク接続またはGitHub認証を確認してください。"
    read -p "Press Enter to close..."
    exit 1
fi

# ---- 4. 依存パッケージ更新 ----
echo ""
echo "--- Python パッケージを更新中... ---"
pip3 install -r requirements.txt --quiet
echo "[OK] Python パッケージ完了"

echo ""
echo "--- Node.js パッケージを更新中... ---"
npm install --silent 2>/dev/null
echo "[OK] Node.js パッケージ完了"

# ---- 5. .command に実行権限付与 ----
chmod +x ./*.command 2>/dev/null

# ---- 6. 完了 ----
echo ""
echo "============================================="
echo "  アップデート完了！"
echo "============================================="
echo ""
echo "  config.json の設定はそのまま保持されています。"
echo ""
read -p "Press Enter to close..."

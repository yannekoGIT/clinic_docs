#!/bin/bash
# セットアップ / アップデート
cd "$(dirname "$0")"
echo ""
echo "============================================="
echo "  セットアップ / アップデート"
echo "============================================="
echo ""

# ---- 1. config.json が無ければテンプレからコピー ----
if [ ! -f config.json ]; then
    if [ -f config.example.json ]; then
        cp config.example.json config.json
        echo "[初回] config.json を作成しました。"
        echo "       クリニック情報・API設定を編集してください。"
        echo ""
    else
        echo "[ERROR] config.example.json が見つかりません。"
        read -p "Press Enter to close..."
        exit 1
    fi
else
    echo "[OK] config.json は既存のものを保持します。"
fi

# ---- 2. Python チェック ----
if command -v python3 &>/dev/null; then
    echo "[OK] $(python3 --version)"
else
    echo "[NG] Python3 が見つかりません。"
    echo "    brew install python3 または https://www.python.org/downloads/"
    read -p "Press Enter to close..."
    exit 1
fi

# ---- 3. Node.js チェック ----
if command -v node &>/dev/null; then
    echo "[OK] Node.js $(node --version)"
else
    echo "[NG] Node.js が見つかりません。"
    echo "    brew install node または https://nodejs.org/"
    read -p "Press Enter to close..."
    exit 1
fi

# ---- 4. 依存パッケージインストール ----
echo ""
echo "--- Python パッケージをインストール中... ---"
pip3 install -r requirements.txt --quiet
echo "[OK] Python パッケージ完了"

echo ""
echo "--- Node.js パッケージをインストール中... ---"
npm install --silent 2>/dev/null
echo "[OK] Node.js パッケージ完了"

# ---- 5. 出力フォルダ作成 ----
mkdir -p output

# ---- 6. .command ファイルに実行権限付与 ----
chmod +x ./*.command 2>/dev/null

# ---- 7. 完了 ----
echo ""
echo "============================================="
echo "  セットアップ完了！"
echo "============================================="
echo ""
echo "  初回の場合: config.json を編集してください。"
echo "  アップデートの場合: 設定はそのまま保持されています。"
echo ""
read -p "Press Enter to close..."

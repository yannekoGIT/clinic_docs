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
        echo "       クリニック情報を編集してください。"
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
    echo ""
    echo "  ***********************************************"
    echo "  *  Python3 が見つかりません                    *"
    echo "  ***********************************************"
    echo ""
    echo "  以下のいずれかでインストールしてください:"
    echo "    brew install python3"
    echo "    または https://www.python.org/downloads/"
    echo ""
    read -p "Press Enter to close..."
    exit 1
fi

# ---- 3. Node.js チェック ----
if command -v node &>/dev/null; then
    echo "[OK] Node.js $(node --version)"
else
    echo ""
    echo "  ***********************************************"
    echo "  *  Node.js が見つかりません                    *"
    echo "  ***********************************************"
    echo ""
    echo "  以下のいずれかでインストールしてください:"
    echo "    brew install node"
    echo "    または https://nodejs.org/"
    echo ""
    read -p "Press Enter to close..."
    exit 1
fi

# ---- 4. Codex CLI チェック ----
if command -v codex &>/dev/null; then
    echo "[OK] Codex CLI $(codex --version 2>&1)"
else
    echo ""
    echo "  --- Codex CLI をインストール中... ---"
    npm install -g @openai/codex 2>/dev/null
    if [ $? -ne 0 ]; then
        echo ""
        echo "  ***********************************************"
        echo "  *  Codex CLI のインストールに失敗しました      *"
        echo "  ***********************************************"
        echo ""
        echo "  以下を手動で試してください:"
        echo "    npm install -g @openai/codex"
        echo ""
        echo "  インストール後、このスクリプトを再実行してください。"
        echo ""
        read -p "Press Enter to close..."
        exit 1
    fi
    echo "  [OK] Codex CLI インストール完了"
fi

# ---- 5. Codex CLI 認証チェック ----
if [ -n "$OPENAI_API_KEY" ]; then
    echo "[OK] Codex 認証済み"
elif [ -f "$HOME/.codex/config.toml" ]; then
    echo "[OK] Codex 認証済み"
else
    echo ""
    echo "  ***********************************************"
    echo "  *  Codex の認証設定が必要です（初回のみ）      *"
    echo "  ***********************************************"
    echo ""
    echo "  ブラウザが開くので、OpenAI アカウントでログインしてください。"
    echo "  （アカウントが無い場合はその場で作成できます）"
    echo ""
    echo "  ※ APIの利用には支払い設定が必要です（従量課金）"
    echo "    https://platform.openai.com/settings/organization/billing"
    echo ""

    codex auth
    if [ $? -eq 0 ]; then
        echo ""
        echo "  [OK] Codex 認証完了"
    else
        echo ""
        echo "  [注意] 認証がスキップされました。"
        echo "         書類生成の実行時に再度認証を求められます。"
    fi
fi

# ---- 6. 依存パッケージインストール ----
echo ""
echo "--- Python パッケージをインストール中... ---"
pip3 install -r requirements.txt --quiet
echo "[OK] Python パッケージ完了"

echo ""
echo "--- Node.js パッケージをインストール中... ---"
npm install --silent 2>/dev/null
echo "[OK] Node.js パッケージ完了"

# ---- 7. 出力フォルダ作成 ----
mkdir -p output

# ---- 8. .command ファイルに実行権限付与 ----
chmod +x ./*.command 2>/dev/null

# ---- 9. 完了 ----
echo ""
echo "============================================="
echo "  セットアップ完了！"
echo "============================================="
echo ""
echo "  初回の場合:"
echo "    1. config.json のクリニック情報を編集してください"
echo "    2. input フォルダにカルテ(.txt)を配置して書類生成を実行"
echo ""
echo "  アップデートの場合:"
echo "    設定はそのまま保持されています。"
echo ""
read -p "Press Enter to close..."

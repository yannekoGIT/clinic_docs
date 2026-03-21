#!/bin/bash
# 紹介状（診療情報提供書）作成 — コピペ用テキスト(.txt)を生成
cd "$(dirname "$0")"
echo ""
echo "============================================="
echo "  紹介状（診療情報提供書）作成"
echo "  コピペ用テキスト（.txt）を生成します"
echo "============================================="
echo ""

if ! command -v python3 &> /dev/null; then
    echo "[ERROR] Python3 が見つかりません。"
    echo "  brew install python3 でインストールしてください。"
    read -p "Press Enter to close..."
    exit 1
fi

python3 run_batch.py referral "$@"
echo ""
if [ $? -eq 0 ]; then
    echo "出力先: output フォルダを確認してください。"
else
    echo "エラーが発生しました。"
fi
echo ""
read -p "Press Enter to close..."

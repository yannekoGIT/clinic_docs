#!/bin/bash
# 障害年金診断書（精神の障害用）作成 — Excel(.xlsx)
cd "$(dirname "$0")"
echo ""
echo "============================================="
echo "  障害年金診断書（精神の障害用）作成"
echo "============================================="
echo ""

if ! command -v python3 &> /dev/null; then
    echo "[ERROR] Python3 が見つかりません。"
    read -p "Press Enter to close..."
    exit 1
fi

python3 run_batch.py shougai_nenkin "$@"
echo ""
if [ $? -eq 0 ]; then
    echo "出力先: output フォルダを確認してください。"
else
    echo "エラーが発生しました。"
fi
echo ""
read -p "Press Enter to close..."

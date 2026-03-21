#!/bin/bash
# 自立支援医療診断書（精神通院医療用）作成 — Word版(.docx)
cd "$(dirname "$0")"
echo ""
echo "============================================="
echo "  自立支援医療診断書（精神通院医療用）作成"
echo "  Word版（.docx）で生成します"
echo "============================================="
echo ""

if ! command -v python3 &> /dev/null; then
    echo "[ERROR] Python3 が見つかりません。"
    read -p "Press Enter to close..."
    exit 1
fi

python3 run_batch.py jiritsu_shien_word "$@"
echo ""
if [ $? -eq 0 ]; then
    echo "出力先: output フォルダを確認してください。"
else
    echo "エラーが発生しました。"
fi
echo ""
read -p "Press Enter to close..."

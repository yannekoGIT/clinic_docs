#!/bin/bash
# 環境チェック
cd "$(dirname "$0")"
echo ""
echo "============================================="
echo "  環境チェック"
echo "============================================="
echo ""
python3 generate.py --check
echo ""
read -p "Press Enter to close..."

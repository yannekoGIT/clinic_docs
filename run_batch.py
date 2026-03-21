#!/usr/bin/env python3
"""
汎用バッチ生成スクリプト

指定された書類タイプの input/{doc_type}/ フォルダ内の .txt ファイルを処理し、
output/YYYY-MM-DD/ に書類(.docx)を生成する。

使い方:
  python run_batch.py referral
  python run_batch.py jiritsu_shien
  python run_batch.py shougai_nenkin
  python run_batch.py jiritsu_shien --from-json

オプション:
  --provider api|local   LLMプロバイダーを指定（デフォルト: config.jsonの設定）
  --json-only            LLMによるJSON変換のみ（docx生成しない）
  --from-json            input/{doc_type}/*.json から直接docx生成（LLMスキップ）
  --save-json            中間JSONも output/ に保存
  --config PATH          設定ファイルパス（デフォルト: config.json）
"""
import argparse
import json
import os
import sys
from datetime import datetime
from pathlib import Path

ROOT = Path(__file__).parent
sys.path.insert(0, str(ROOT))

from generate import DOC_TYPES, karte_to_json, json_to_document, make_output_path
from lib.llm_client import load_config, CodexNotFoundError, CodexAuthError


def find_input_files(input_dir, from_json=False):
    """入力ファイルを検索"""
    ext = ".json" if from_json else ".txt"
    files = sorted(input_dir.glob(f"*{ext}"))
    return [f for f in files if f.name != "README.txt"]


def get_output_ext(doc_type):
    """書類タイプに応じた出力拡張子を返す"""
    return DOC_TYPES.get(doc_type, {}).get("output_ext", ".docx")


def generate_output_path(doc_type, input_file, ext=None):
    """入力ファイル名から日付フォルダ付き出力パスを生成
    通常: output/YYYY-MM-DD/{doc_type}_{stem}_{HHMMSS}.ext
    JSON: output/YYYY-MM-DD/json/{doc_type}_{stem}_{HHMMSS}.json
    """
    if ext is None:
        ext = get_output_ext(doc_type)
    now = datetime.now()
    date_dir = now.strftime("%Y-%m-%d")
    timestamp = now.strftime("%H%M%S")
    if ext == ".json":
        out_dir = ROOT / "output" / date_dir / "json"
    else:
        out_dir = ROOT / "output" / date_dir
    out_dir.mkdir(parents=True, exist_ok=True)
    return out_dir / f"{doc_type}_{input_file.stem}_{timestamp}{ext}"


def process_txt_file(doc_type, input_file, config, save_json=False, json_only=False):
    """テキストファイルからLLM→JSON(→docx)を生成"""
    print(f"\n--- 処理中: {input_file.name} ---")

    with open(input_file, "r", encoding="utf-8") as f:
        karte_text = f.read()

    if not karte_text.strip():
        print(f"  [SKIP] ファイルが空です: {input_file.name}")
        return False

    # LLMでJSON変換
    try:
        json_data = karte_to_json(doc_type, karte_text, config)
    except CodexNotFoundError as e:
        print(f"  [ERROR] {e}", file=sys.stderr)
        return False
    except CodexAuthError as e:
        print(f"  [ERROR] {e}", file=sys.stderr)
        return False
    except Exception as e:
        print(f"  [ERROR] LLM変換失敗: {e}")
        return False

    print("  [OK] JSON変換完了")

    # JSON保存
    if save_json or json_only:
        json_path = generate_output_path(doc_type, input_file, ".json")
        with open(json_path, "w", encoding="utf-8") as f:
            json.dump(json_data, f, ensure_ascii=False, indent=2)
        print(f"  [OK] JSON保存: {json_path.name}")

    if json_only:
        return True

    # 書類生成
    output_path = generate_output_path(doc_type, input_file)
    success = json_to_document(doc_type, json_data, output_path)
    if success:
        print(f"  [OK] 書類生成: {output_path.name}")
    return success


def process_json_file(doc_type, input_file):
    """JSONファイルから直接書類生成"""
    print(f"\n--- 処理中: {input_file.name} ---")

    try:
        with open(input_file, "r", encoding="utf-8") as f:
            json_data = json.load(f)
    except (json.JSONDecodeError, OSError) as e:
        print(f"  [ERROR] JSON読み込み失敗: {e}")
        return False

    output_path = generate_output_path(doc_type, input_file)
    success = json_to_document(doc_type, json_data, output_path)
    if success:
        print(f"  [OK] 書類生成: {output_path.name}")
    return success


def main():
    parser = argparse.ArgumentParser(description="書類バッチ生成")
    parser.add_argument("doc_type", choices=DOC_TYPES.keys(),
                        help="書類タイプ: " + ", ".join(
                            f"{k}({v['name']})" for k, v in DOC_TYPES.items()))
    parser.add_argument("--provider", "-p", choices=["codex", "api", "local"],
                        help="LLMプロバイダー（codex/api/local）")
    parser.add_argument("--json-only", action="store_true",
                        help="JSON変換のみ（docx生成しない）")
    parser.add_argument("--from-json", action="store_true",
                        help="JSONから直接docx生成（LLMスキップ）")
    parser.add_argument("--save-json", action="store_true",
                        help="中間JSONも保存")
    parser.add_argument("--config", "-c", default=str(ROOT / "config.json"),
                        help="設定ファイルパス")
    args = parser.parse_args()

    doc_type = args.doc_type
    type_info = DOC_TYPES[doc_type]
    # 派生タイプは元タイプと同じ入力フォルダを使用
    input_folder = doc_type.replace("_word", "").replace("_docx", "")
    input_dir = ROOT / "input" / input_folder

    # フォルダ準備
    input_dir.mkdir(parents=True, exist_ok=True)

    # 設定読み込み
    config = load_config(args.config)
    if args.provider:
        config["llm"]["provider"] = args.provider

    print("=" * 50)
    print(f"  {type_info['name']} バッチ生成")
    print("=" * 50)
    print(f"入力フォルダ: {input_dir}")
    print(f"出力フォルダ: {ROOT / 'output' / datetime.now().strftime('%Y-%m-%d')}")

    # 入力ファイル検索
    files = find_input_files(input_dir, from_json=args.from_json)

    if not files:
        ext = ".json" if args.from_json else ".txt"
        print(f"\n[INFO] 入力ファイルがありません。")
        print(f"  {input_dir} に {ext} ファイルを配置してください。")
        if (input_dir / "sample_karte.txt").exists():
            print(f"  サンプル: input/{doc_type}/sample_karte.txt を参照")
        return

    print(f"\n対象ファイル: {len(files)}件")
    for f in files:
        print(f"  - {f.name}")

    # 処理
    success_count = 0
    fail_count = 0

    for input_file in files:
        if args.from_json:
            ok = process_json_file(doc_type, input_file)
        else:
            ok = process_txt_file(doc_type, input_file, config,
                                  save_json=args.save_json,
                                  json_only=args.json_only)
        if ok:
            success_count += 1
        else:
            fail_count += 1

    # 結果サマリ
    print("\n" + "=" * 50)
    print(f"  完了: {success_count}件成功 / {fail_count}件失敗")
    print("=" * 50)

    if fail_count > 0:
        sys.exit(1)


if __name__ == "__main__":
    main()

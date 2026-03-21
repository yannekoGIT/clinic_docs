#!/usr/bin/env python3
"""
クリニック書類自動生成ツール
カルテ情報 → LLM(JSON) → 書類生成（docx / xls / xlsx）

Usage:
  python generate.py referral       "カルテのテキスト..."
  python generate.py diagnosis      "カルテのテキスト..."
  python generate.py jiritsu_shien  "カルテのテキスト..."
  python generate.py shougai_nenkin "カルテのテキスト..."

  python generate.py referral       --file karte.txt
  python generate.py referral       --json data.json     # LLMスキップ、JSONから直接生成
  python generate.py --provider codex referral "..."      # Codex CLI使用（デフォルト）
  python generate.py --provider api   referral "..."      # OpenAI互換API使用
  python generate.py --provider local referral "..."      # ローカルLLM使用
"""
import argparse
import json
import os
import shutil
import subprocess
import sys
from datetime import datetime
from pathlib import Path

# プロジェクトルート
ROOT = Path(__file__).parent
LIB = ROOT / "lib"
sys.path.insert(0, str(ROOT))

from lib.llm_client import load_config, call_llm, extract_json, CodexNotFoundError, CodexAuthError


def validate_json(data, schema_path):
    """JSONデータをスキーマの必須フィールド・型でバリデーション（標準ライブラリのみ）"""
    with open(schema_path, "r", encoding="utf-8") as f:
        schema = json.load(f)

    errors = []
    required = schema.get("required", [])
    properties = schema.get("properties", {})

    for field in required:
        if field not in data or data[field] is None:
            errors.append(f"必須フィールド '{field}' がありません")
        elif field in properties:
            expected = properties[field].get("type")
            if expected == "string" and not isinstance(data[field], str):
                errors.append(f"'{field}' は文字列である必要があります（実際: {type(data[field]).__name__}）")
            elif expected == "array" and not isinstance(data[field], list):
                errors.append(f"'{field}' は配列である必要があります（実際: {type(data[field]).__name__}）")
            elif expected == "object" and not isinstance(data[field], dict):
                errors.append(f"'{field}' はオブジェクトである必要があります（実際: {type(data[field]).__name__}）")

    # patient内の必須フィールドをチェック
    if "patient" in properties and isinstance(data.get("patient"), dict):
        patient_schema = properties["patient"]
        for pf in patient_schema.get("required", []):
            if pf not in data["patient"] or data["patient"][pf] is None:
                errors.append(f"patient.{pf} がありません")

    return errors


# 書類タイプの定義（パスは lib/ 配下）
# renderer: "nodejs" = Node.js/docx-js で docx 生成
#           "python" = Python で公式テンプレート(xls/xlsx)に書き込み
DOC_TYPES = {
    "referral": {
        "name": "紹介状（診療情報提供書）",
        "prompt_file": "lib/prompts/referral.txt",
        "schema_file": "lib/schemas/referral.json",
        "renderer": "direct_text",
        "output_ext": ".txt",
    },
    "diagnosis": {
        "name": "診断書",
        "prompt_file": "lib/prompts/diagnosis_certificate.txt",
        "template_js": "lib/templates/diagnosis_certificate.js",
        "schema_file": "lib/schemas/diagnosis_certificate.json",
        "renderer": "nodejs",
        "output_ext": ".docx",
    },
    "jiritsu_shien": {
        "name": "自立支援医療診断書（精神通院医療用）",
        "prompt_file": "lib/prompts/jiritsu_shien.txt",
        "template_py": "lib/templates/jiritsu_shien.py",
        "schema_file": "lib/schemas/jiritsu_shien.json",
        "renderer": "python",
        "output_ext": ".xls",
    },
    "jiritsu_shien_word": {
        "name": "自立支援医療診断書（精神通院医療用）Word版",
        "prompt_file": "lib/prompts/jiritsu_shien.txt",
        "template_py": "lib/templates/jiritsu_shien_word.py",
        "schema_file": "lib/schemas/jiritsu_shien.json",
        "renderer": "python",
        "output_ext": ".docx",
    },
    "shougai_nenkin": {
        "name": "障害年金診断書（精神の障害用・様式第120号の4）",
        "prompt_file": "lib/prompts/shougai_nenkin.txt",
        "template_py": "lib/templates/shougai_nenkin.py",
        "schema_file": "lib/schemas/shougai_nenkin.json",
        "renderer": "python",
        "output_ext": ".xlsx",
    },
}


def get_clinic_info(config):
    """設定からクリニック情報をテキスト化"""
    c = config.get("clinic", {})
    lines = []
    if c.get("name"):
        lines.append(f"医療機関名: {c['name']}")
    if c.get("department"):
        lines.append(f"診療科: {c['department']}")
    if c.get("doctor"):
        lines.append(f"医師名: {c['doctor']}")
    if c.get("address"):
        lines.append(f"住所: {c['address']}")
    if c.get("phone"):
        lines.append(f"電話番号: {c['phone']}")
    return "\n".join(lines) if lines else "（未設定）"


def make_output_path(doc_type, base_dir=None, ext=".docx"):
    """日付フォルダ付きの出力パスを生成
    通常: output/YYYY-MM-DD/{doc_type}_{HHMMSS}.docx
    JSON: output/YYYY-MM-DD/json/{doc_type}_{HHMMSS}.json
    """
    now = datetime.now()
    date_dir = now.strftime("%Y-%m-%d")
    timestamp = now.strftime("%H%M%S")
    if ext == ".json":
        out_dir = (base_dir or ROOT / "output") / date_dir / "json"
    else:
        out_dir = (base_dir or ROOT / "output") / date_dir
    out_dir.mkdir(parents=True, exist_ok=True)
    return out_dir / f"{doc_type}_{timestamp}{ext}"


def karte_to_json(doc_type, karte_text, config):
    """カルテテキストをLLMでJSONに変換"""
    type_info = DOC_TYPES[doc_type]

    # プロンプト読み込み
    prompt_path = ROOT / type_info["prompt_file"]
    with open(prompt_path, "r", encoding="utf-8") as f:
        prompt_template = f.read()

    # クリニック情報を埋め込み
    clinic_info = get_clinic_info(config)
    prompt = prompt_template.replace("{clinic_info}", clinic_info).replace("{karte_text}", karte_text)

    # 文字数制限を埋め込み（config.json の char_limits から）
    char_limits_text = _build_char_limits_text(doc_type, config)
    prompt = prompt.replace("{char_limits}", char_limits_text)

    # カスタムプロンプトを追加（config.json の custom_prompt から）
    custom_key = doc_type.replace("_word", "").replace("_docx", "")
    custom = config.get("custom_prompt", {}).get(custom_key, "")
    if custom:
        custom_block = f"\n## 追加指示\n{custom}\n"
    else:
        custom_block = ""
    # {custom_instructions} プレースホルダーがあればそこに挿入、なければ末尾に追加
    if "{custom_instructions}" in prompt:
        prompt = prompt.replace("{custom_instructions}", custom_block)
    elif custom_block:
        prompt += "\n" + custom_block

    # direct_text モード: LLM出力をそのままテキストとして返す（JSON変換しない）
    renderer = type_info.get("renderer", "")
    if renderer == "direct_text":
        print(f"[INFO] LLMにカルテ情報を送信中... ({config['llm']['provider']})")
        response = call_llm(prompt, config=config)
        return {"_raw_text": response.strip()}

    # JSON出力モード: Codex CLI使用時はスキーマを渡して構造化出力を強制
    schema_path = ROOT / type_info["schema_file"]
    print(f"[INFO] LLMにカルテ情報を送信中... ({config['llm']['provider']})")
    response = call_llm(prompt, config=config,
                        schema_path=schema_path if schema_path.exists() else None)
    data = extract_json(response)

    # 日付が未設定なら今日の日付
    if not data.get("date"):
        data["date"] = datetime.now().strftime("%Y-%m-%d")

    return data


def _build_char_limits_text(doc_type, config):
    """config.json の char_limits からプロンプト用テキストを生成"""
    # jiritsu_shien_word は jiritsu_shien の設定を使う
    key = doc_type.replace("_word", "")
    limits = config.get("char_limits", {}).get(key, {})
    if not limits:
        return "（文字数制限の設定なし）"
    lines = []
    for field, limit in limits.items():
        # _始まりは説明用キーなのでスキップ
        if field.startswith("_"):
            continue
        lines.append(f"- **{field}**: {limit}文字以内")
    return "\n".join(lines)


def check_node_available():
    """Node.jsが利用可能か検査"""
    if shutil.which("node") is None:
        print(
            "[ERROR] Node.js が見つかりません。\n"
            "\n"
            "docx生成には Node.js が必要です。\n"
            "=== セットアップ手順 ===\n"
            "1. https://nodejs.org/ から Node.js (v18以上) をインストール\n"
            "2. インストール後、このフォルダで以下を実行:\n"
            "     npm install\n"
            "3. 再度実行してください。",
            file=sys.stderr,
        )
        sys.exit(1)


def json_to_document(doc_type, json_data, output_path):
    """JSONから書類を生成（書類タイプに応じてNode.js、Python、またはテキスト直接出力）"""
    type_info = DOC_TYPES[doc_type]
    renderer = type_info.get("renderer", "nodejs")

    if renderer == "direct_text":
        return _save_direct_text(json_data, output_path)
    elif renderer == "python":
        return _json_to_template_py(doc_type, json_data, output_path)
    else:
        return _json_to_docx_nodejs(doc_type, json_data, output_path)


def _save_direct_text(data, output_path):
    """LLMの生テキスト出力をそのままファイルに保存"""
    text = data.get("_raw_text", "")
    if not text:
        print("[ERROR] テキスト出力が空です", file=sys.stderr)
        return False
    output_path = Path(output_path)
    output_path.parent.mkdir(parents=True, exist_ok=True)
    with open(output_path, "w", encoding="utf-8") as f:
        f.write(text)
    return True


def _json_to_template_py(doc_type, json_data, output_path):
    """JSONからPythonスクリプトで公式テンプレート(xls/xlsx)に書き込み"""
    type_info = DOC_TYPES[doc_type]
    template_py = ROOT / type_info["template_py"]

    if not template_py.exists():
        print(f"[ERROR] テンプレートスクリプトが見つかりません: {template_py}", file=sys.stderr)
        return False

    # 一時JSONファイル
    output_path = Path(output_path)
    output_path.parent.mkdir(parents=True, exist_ok=True)
    tmp_json = output_path.parent / f"_tmp_{doc_type}.json"
    with open(tmp_json, "w", encoding="utf-8") as f:
        json.dump(json_data, f, ensure_ascii=False, indent=2)

    try:
        result = subprocess.run(
            [sys.executable, str(template_py), str(tmp_json), str(output_path)],
            cwd=str(ROOT),
            capture_output=True,
            text=True,
            encoding="utf-8",
            errors="replace",
        )
        if result.returncode != 0:
            print(f"[ERROR] テンプレート書き込みエラー:\n{result.stderr}", file=sys.stderr)
            return False
        if result.stdout:
            _safe_print(result.stdout.strip())
        return True
    finally:
        try:
            if tmp_json.exists():
                tmp_json.unlink()
        except OSError:
            pass


def _safe_print(text):
    """コンソールのエンコーディング(cp932等)で出力できない文字を安全に処理"""
    try:
        print(text)
    except UnicodeEncodeError:
        print(text.encode(sys.stdout.encoding or "utf-8", errors="replace").decode(sys.stdout.encoding or "utf-8", errors="replace"))


def _json_to_docx_nodejs(doc_type, json_data, output_path):
    """JSONからNode.js/docx-jsでdocxを生成"""
    check_node_available()

    # node_modules が無ければ npm install を案内
    if not (ROOT / "node_modules" / "docx").exists():
        print(
            "[ERROR] Node.js依存パッケージが未インストールです。\n"
            "以下を実行してください:\n"
            f"  cd {ROOT}\n"
            "  npm install",
            file=sys.stderr,
        )
        return False

    type_info = DOC_TYPES[doc_type]
    template_js = ROOT / type_info["template_js"]

    # 一時JSONファイル（出力先と同じ場所に作成）
    output_path = Path(output_path)
    output_path.parent.mkdir(parents=True, exist_ok=True)
    tmp_json = output_path.parent / f"_tmp_{doc_type}.json"
    with open(tmp_json, "w", encoding="utf-8") as f:
        json.dump(json_data, f, ensure_ascii=False, indent=2)

    try:
        result = subprocess.run(
            ["node", str(template_js), str(tmp_json), str(output_path)],
            cwd=str(ROOT),
            capture_output=True,
            text=True,
            encoding="utf-8",
            errors="replace",
        )
        if result.returncode != 0:
            print(f"[ERROR] docx生成エラー:\n{result.stderr}", file=sys.stderr)
            return False
        if result.stdout:
            print(result.stdout.strip())
        return True
    finally:
        try:
            if tmp_json.exists():
                tmp_json.unlink()
        except OSError:
            pass  # 一時ファイル削除失敗は無視


# 後方互換性のためのエイリアス
json_to_docx = json_to_document


def run_setup_check(config_path):
    """環境セットアップを検証し、結果を表示"""
    print("=== 環境セットアップ検証 ===\n")
    ok = True

    # 1. config.json
    try:
        config = load_config(config_path)
        print("[OK] config.json 読み込み成功")
    except Exception as e:
        print(f"[NG] config.json 読み込み失敗: {e}")
        return

    # 2. クリニック情報
    clinic = config.get("clinic", {})
    empty_fields = [k for k, v in clinic.items() if not v]
    if empty_fields:
        print(f"[WARN] config.json のクリニック情報が未設定: {', '.join(empty_fields)}")
        print("       → config.json の clinic セクションを編集してください")
        ok = False
    else:
        print("[OK] クリニック情報 設定済み")

    # 3. Node.js
    if shutil.which("node"):
        result = subprocess.run(["node", "--version"], capture_output=True, text=True)
        print(f"[OK] Node.js {result.stdout.strip()}")
    else:
        print("[NG] Node.js が見つかりません → https://nodejs.org/ からインストール")
        ok = False

    # 4. npm dependencies
    if (ROOT / "node_modules" / "docx").exists():
        print("[OK] npm依存パッケージ (docx) インストール済み")
    else:
        print("[NG] npm依存パッケージ未インストール → npm install を実行してください")
        ok = False

    # 5. LLMプロバイダー
    provider = config["llm"]["provider"]
    print(f"\n--- LLMプロバイダー: {provider} ---")

    if provider == "codex":
        from lib.llm_client import check_codex_available, check_codex_auth
        try:
            check_codex_available()
            print("[OK] Codex CLI インストール済み")
        except CodexNotFoundError as e:
            print(f"[NG] {e}")
            ok = False
        try:
            check_codex_auth()
            print("[OK] Codex CLI 認証設定済み")
        except CodexAuthError as e:
            print(f"[NG] {e}")
            ok = False
    else:
        from lib.llm_client import get_llm_settings
        settings = get_llm_settings(config)
        print(f"  Base URL: {settings.get('base_url', 'N/A')}")
        if settings.get("api_key"):
            print("[OK] APIキー 設定済み")
        elif provider == "local":
            print("[INFO] ローカルLLM — APIキー不要（サーバー起動を確認してください）")
        else:
            print("[NG] APIキーが未設定 → 環境変数を確認してください")
            ok = False

    # 6. テンプレートファイル
    missing = []
    for key, info in DOC_TYPES.items():
        template_key = "template_py" if info.get("renderer") == "python" else "template_js"
        for fkey in ("prompt_file", template_key, "schema_file"):
            p = ROOT / info[fkey]
            if not p.exists():
                missing.append(str(p))
    if missing:
        print(f"[NG] テンプレートファイル欠損: {', '.join(missing)}")
        ok = False
    else:
        print("[OK] テンプレートファイル 全て存在")

    print()
    if ok:
        print("全チェック通過 — 利用可能です。")
    else:
        print("上記の [NG] / [WARN] 項目を解決してください。")


def main():
    parser = argparse.ArgumentParser(description="クリニック書類自動生成ツール")
    parser.add_argument("doc_type", nargs="?", choices=DOC_TYPES.keys(), default=None,
                        help="書類タイプ: referral(紹介状), diagnosis(診断書), jiritsu_shien(自立支援), shougai_nenkin(障害年金)")
    parser.add_argument("karte_text", nargs="?", default=None,
                        help="カルテテキスト（直接入力）")
    parser.add_argument("--file", "-f", help="カルテテキストファイル")
    parser.add_argument("--json", "-j", help="JSONファイル（LLMをスキップして直接docx生成）")
    parser.add_argument("--output", "-o", help="出力ファイルパス")
    parser.add_argument("--provider", "-p", choices=["codex", "api", "local"],
                        help="LLMプロバイダー（codex/api/local）")
    parser.add_argument("--config", "-c", default=str(ROOT / "config.json"),
                        help="設定ファイルパス")
    parser.add_argument("--save-json", action="store_true",
                        help="中間JSONファイルも保存する")
    parser.add_argument("--preview", action="store_true",
                        help="JSON生成後にプレビュー表示し、docx生成前に確認する")
    parser.add_argument("--check", action="store_true",
                        help="環境セットアップを検証して終了（書類生成しない）")

    args = parser.parse_args()

    # --check: 環境検証モード
    if args.check:
        run_setup_check(args.config)
        sys.exit(0)

    # doc_type は --check 以外では必須
    if not args.doc_type:
        parser.error("書類タイプ（referral / diagnosis / jiritsu_shien / shougai_nenkin）を指定してください")

    # 設定読み込み（config.json が無ければ config.example.json からコピー）
    config_path = Path(args.config)
    if not config_path.exists():
        example_path = config_path.parent / "config.example.json"
        if example_path.exists():
            shutil.copy2(example_path, config_path)
            print(f"[INFO] {config_path.name} を config.example.json から作成しました。")
            print(f"       クリニック情報を編集してください: {config_path}")
        else:
            print(f"[ERROR] 設定ファイルが見つかりません: {config_path}", file=sys.stderr)
            sys.exit(1)

    config = load_config(args.config)
    if args.provider:
        config["llm"]["provider"] = args.provider

    type_info = DOC_TYPES[args.doc_type]

    # 出力パス（--output指定なしなら日付フォルダ付き自動生成）
    output_ext = type_info.get("output_ext", ".docx")
    if args.output:
        output_path = Path(args.output)
    else:
        output_path = make_output_path(args.doc_type, ext=output_ext)

    output_path.parent.mkdir(parents=True, exist_ok=True)

    print(f"=== {type_info['name']} 生成 ===")

    # JSON直接指定の場合
    if args.json:
        with open(args.json, "r", encoding="utf-8") as f:
            json_data = json.load(f)
        print(f"[INFO] JSONファイルから読み込み: {args.json}")

    else:
        # カルテテキスト取得
        if args.file:
            with open(args.file, "r", encoding="utf-8") as f:
                karte_text = f.read()
            print(f"[INFO] カルテファイル読み込み: {args.file}")
        elif args.karte_text:
            karte_text = args.karte_text
        else:
            eof_hint = "Ctrl+Z → Enter" if sys.platform == "win32" else "Ctrl+D"
            print(f"[INFO] カルテ情報を入力してください（{eof_hint}で終了）:")
            karte_text = sys.stdin.read()

        if not karte_text.strip():
            print("[ERROR] カルテ情報が空です", file=sys.stderr)
            sys.exit(1)

        # LLMでJSON変換
        try:
            json_data = karte_to_json(args.doc_type, karte_text, config)
        except CodexNotFoundError as e:
            print(f"\n[ERROR] {e}", file=sys.stderr)
            sys.exit(2)
        except CodexAuthError as e:
            print(f"\n[ERROR] {e}", file=sys.stderr)
            sys.exit(3)
        print("[INFO] LLM処理完了")

        # 中間JSON保存（direct_textモード以外、jsonサブフォルダに格納）
        if args.save_json and "_raw_text" not in json_data:
            json_path = make_output_path(args.doc_type, ext=".json")
            with open(json_path, "w", encoding="utf-8") as f:
                json.dump(json_data, f, ensure_ascii=False, indent=2)
            print(f"[INFO] JSON保存: {json_path}")

    # スキーマバリデーション（direct_textモードではスキップ）
    renderer = type_info.get("renderer", "")
    if renderer != "direct_text":
        schema_path = ROOT / type_info["schema_file"]
        if schema_path.exists():
            errors = validate_json(json_data, schema_path)
            if errors:
                print("[WARN] JSONバリデーション警告:", file=sys.stderr)
                for e in errors:
                    print(f"  - {e}", file=sys.stderr)
                print("[INFO] 警告がありますが、生成を続行します")
            else:
                print("[INFO] JSONバリデーション OK")

    # プレビュー表示
    if args.preview:
        print("\n--- JSON プレビュー ---")
        print(json.dumps(json_data, ensure_ascii=False, indent=2))
        print("--- プレビュー終了 ---\n")
        answer = input("docxを生成しますか？ [Y/n]: ").strip().lower()
        if answer in ("n", "no"):
            # プレビューのみの場合でもJSONは保存
            json_path = output_path.with_suffix(".json")
            with open(json_path, "w", encoding="utf-8") as f:
                json.dump(json_data, f, ensure_ascii=False, indent=2)
            print(f"[INFO] JSON保存: {json_path}")
            print("[INFO] docx生成をスキップしました")
            sys.exit(0)

    # 書類生成
    success = json_to_document(args.doc_type, json_data, output_path)
    if success:
        print(f"\n[OK] {type_info['name']}を生成しました: {output_path}")
    else:
        print(f"\n[FAIL] 生成に失敗しました", file=sys.stderr)
        sys.exit(1)


if __name__ == "__main__":
    main()

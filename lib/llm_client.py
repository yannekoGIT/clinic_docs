"""
LLM抽象化レイヤー
Codex CLI（デフォルト）、ローカル（Ollama等）、API（OpenAI互換）を切り替え可能
"""
import json
import os
import shutil
import subprocess
import sys
import urllib.request
import urllib.error


def load_config(config_path="config.json"):
    with open(config_path, "r", encoding="utf-8") as f:
        return json.load(f)


def get_llm_settings(config):
    """現在のプロバイダー設定を返す"""
    provider = config["llm"]["provider"]
    if provider == "codex":
        codex_cfg = config["llm"].get("codex", {})
        return {
            "provider": "codex",
            "model": codex_cfg.get("model", "o4-mini"),
        }
    settings = config["llm"][provider]
    return {
        "provider": provider,
        "base_url": settings["base_url"],
        "api_key": os.environ.get(settings.get("api_key_env", ""), "") if settings.get("api_key_env") else "",
        "model": settings["model"],
    }


# ---------------------------------------------------------------------------
# Codex CLI provider
# ---------------------------------------------------------------------------

class CodexNotFoundError(RuntimeError):
    """Codex CLIがインストールされていない"""
    pass


class CodexAuthError(RuntimeError):
    """Codex CLIの認証が設定されていない"""
    pass


def check_codex_available():
    """Codex CLIが利用可能か検査し、問題があれば明確なエラーを出す"""
    codex_path = shutil.which("codex")
    if codex_path is None:
        raise CodexNotFoundError(
            "Codex CLIが見つかりません。\n"
            "\n"
            "=== セットアップ手順 ===\n"
            "1. Node.js (v22以上) をインストールしてください\n"
            "2. 以下のコマンドでCodex CLIをインストール:\n"
            "     npm install -g @openai/codex\n"
            "\n"
            "インストール後、再度実行してください。"
        )
    return codex_path


def check_codex_auth():
    """Codex CLIの認証が設定されているか検査（ベストエフォート）

    OPENAI_API_KEY または ~/.codex/config.toml での認証を確認する。
    Codex CLI自身が認証を管理するため、ここでは明確に未設定の場合のみ警告する。
    """
    api_key = os.environ.get("OPENAI_API_KEY", "")
    if api_key:
        return  # 環境変数で認証済み

    # Codex CLI自身の認証設定を確認
    codex_config = os.path.expanduser("~/.codex/config.toml")
    if os.path.exists(codex_config):
        return  # Codex CLIの設定ファイルが存在（codex auth 済みの可能性）

    raise CodexAuthError(
        "Codex CLIの認証が設定されていません。\n"
        "\n"
        "=== 認証セットアップ手順 ===\n"
        "以下のいずれかの方法で認証してください:\n"
        "\n"
        "方法1: codex auth で対話的にログイン（推奨）\n"
        "    codex auth\n"
        "\n"
        "方法2: 環境変数を設定\n"
        "  Windows (PowerShell):\n"
        '    $env:OPENAI_API_KEY = "sk-..."\n'
        "  Windows (コマンドプロンプト):\n"
        '    set OPENAI_API_KEY=sk-...\n'
        "\n"
        "APIキーは https://platform.openai.com/api-keys で取得できます。"
    )


def call_codex(prompt, model="o4-mini", schema_path=None):
    """
    Codex CLIを呼び出してテキスト応答を返す

    Args:
        prompt: プロンプト全文
        model: 使用モデル（デフォルト: o4-mini）
        schema_path: JSONスキーマファイルパス（指定時 --output-schema で構造化出力を強制）

    Returns:
        Codex CLIの応答テキスト
    """
    import tempfile

    codex_path = check_codex_available()
    # 認証チェックはベストエフォート — Codex CLI自身が最終判断する
    try:
        check_codex_auth()
    except CodexAuthError:
        # 明確に未設定でも、Codex CLIが独自の認証を持つ可能性があるため続行
        pass

    # 応答を一時ファイルに書き出し、確実に取得する
    tmp_out = None
    try:
        tmp_out = tempfile.NamedTemporaryFile(
            mode="w", suffix=".txt", delete=False, encoding="utf-8"
        )
        tmp_out_path = tmp_out.name
        tmp_out.close()

        cmd = [
            codex_path, "exec",
            "--model", model,
            "-s", "read-only",
            "--skip-git-repo-check",
            "-o", tmp_out_path,
        ]
        if schema_path:
            cmd += ["--output-schema", str(schema_path)]

        # プロンプトは stdin から読み込み（"-" で指定）
        cmd.append("-")

        result = subprocess.run(
            cmd,
            input=prompt,
            capture_output=True,
            text=True,
            encoding="utf-8",
            errors="replace",
            timeout=300,
        )

        if result.returncode != 0:
            stderr = result.stderr.strip()
            # 認証エラーの検出
            if any(kw in stderr.lower() for kw in ["auth", "api key", "unauthorized", "401", "apikey"]):
                raise CodexAuthError(
                    f"Codex CLI認証エラー:\n{stderr}\n\n"
                    "=== 認証の修正方法 ===\n"
                    "1. codex auth  （対話的ログイン、推奨）\n"
                    "2. $env:OPENAI_API_KEY = \"sk-...\"  （PowerShell）\n"
                    "3. set OPENAI_API_KEY=sk-...  （コマンドプロンプト）\n"
                    "\n"
                    "APIキー取得先: https://platform.openai.com/api-keys"
                )
            raise RuntimeError(f"Codex CLIエラー (終了コード {result.returncode}):\n{stderr}")

        # -o で書き出されたファイルを優先、なければ stdout
        output = ""
        try:
            with open(tmp_out_path, "r", encoding="utf-8") as f:
                output = f.read().strip()
        except OSError:
            pass

        if not output:
            output = result.stdout.strip()

        if not output:
            raise RuntimeError("Codex CLIから空の応答が返されました。プロンプトを確認してください。")

        return output

    except subprocess.TimeoutExpired:
        raise RuntimeError(
            "Codex CLIがタイムアウトしました（300秒）。\n"
            "ネットワーク接続を確認してください。"
        )
    finally:
        if tmp_out is not None:
            try:
                os.unlink(tmp_out.name)
            except OSError:
                pass


# ---------------------------------------------------------------------------
# OpenAI互換API provider (既存)
# ---------------------------------------------------------------------------

def call_openai_compatible(prompt, system_prompt="", settings=None):
    """OpenAI互換APIを呼び出してテキスト応答を返す"""
    messages = []
    if system_prompt:
        messages.append({"role": "system", "content": system_prompt})
    messages.append({"role": "user", "content": prompt})

    payload = {
        "model": settings["model"],
        "messages": messages,
        "temperature": 0.1,
        "response_format": {"type": "json_object"},
    }

    url = f"{settings['base_url'].rstrip('/')}/chat/completions"
    headers = {
        "Content-Type": "application/json",
    }
    if settings.get("api_key"):
        headers["Authorization"] = f"Bearer {settings['api_key']}"

    data = json.dumps(payload).encode("utf-8")
    req = urllib.request.Request(url, data=data, headers=headers, method="POST")

    try:
        with urllib.request.urlopen(req, timeout=120) as resp:
            result = json.loads(resp.read().decode("utf-8"))
            return result["choices"][0]["message"]["content"]
    except urllib.error.HTTPError as e:
        body = e.read().decode("utf-8", errors="replace")
        print(f"[LLM Error] HTTP {e.code}: {body}", file=sys.stderr)
        raise
    except urllib.error.URLError as e:
        print(f"[LLM Error] 接続失敗: {e.reason}", file=sys.stderr)
        raise


# ---------------------------------------------------------------------------
# 統合エントリーポイント
# ---------------------------------------------------------------------------

def call_llm(prompt, system_prompt="", config=None, config_path="config.json",
             schema_path=None):
    """
    LLMを呼び出してテキスト応答を返す

    Args:
        prompt: ユーザープロンプト
        system_prompt: システムプロンプト
        config: 設定dict（省略時はファイルから読む）
        config_path: 設定ファイルパス
        schema_path: JSONスキーマパス（Codex CLI使用時に --output-schema で構造化出力を強制）

    Returns:
        LLMの応答テキスト
    """
    if config is None:
        config = load_config(config_path)

    settings = get_llm_settings(config)
    provider = settings["provider"]

    if provider == "codex":
        full_prompt = prompt
        if system_prompt:
            full_prompt = system_prompt + "\n\n" + prompt
        return call_codex(full_prompt, model=settings["model"],
                          schema_path=schema_path)
    else:
        return call_openai_compatible(prompt, system_prompt=system_prompt, settings=settings)


def extract_json(text):
    """LLM応答からJSONを抽出してパースする"""
    import re
    text = text.strip()

    # ```json ... ``` ブロックを抽出（Codex CLIは前後にテキストを付けることがある）
    m = re.search(r"```(?:json)?\s*\n(.*?)\n\s*```", text, re.DOTALL)
    if m:
        text = m.group(1).strip()
    elif text.startswith("```"):
        lines = text.split("\n")
        lines = [l for l in lines if not l.strip().startswith("```")]
        text = "\n".join(lines)

    # JSONオブジェクト部分だけ抽出（前後の説明文を除去）
    if not text.startswith("{"):
        start = text.find("{")
        if start >= 0:
            # 最後の } を探す
            end = text.rfind("}")
            if end > start:
                text = text[start:end + 1]

    return json.loads(text)


if __name__ == "__main__":
    # テスト用
    config = load_config()
    settings = get_llm_settings(config)
    provider = config["llm"]["provider"]
    print(f"Provider: {provider}")
    print(f"Model: {settings['model']}")
    if provider == "codex":
        try:
            check_codex_available()
            print("Codex CLI: インストール済み")
            check_codex_auth()
            print("Codex 認証: 設定済み")
        except CodexNotFoundError as e:
            print(f"Codex CLI: 未インストール\n{e}")
        except CodexAuthError as e:
            print(f"Codex 認証: 未設定\n{e}")
    else:
        print(f"Base URL: {settings.get('base_url', 'N/A')}")
        print(f"API Key set: {'Yes' if settings.get('api_key') else 'No'}")

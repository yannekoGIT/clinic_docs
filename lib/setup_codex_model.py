"""Codex CLI の config.toml にデフォルトモデルを設定するヘルパー"""
import pathlib
import re
import sys

MODEL = "gpt-5.4"

config_path = pathlib.Path.home() / ".codex" / "config.toml"
config_path.parent.mkdir(parents=True, exist_ok=True)

text = config_path.read_text(encoding="utf-8") if config_path.exists() else ""

# 既存の model 行を除去して先頭に挿入
text = re.sub(r"^model\s*=.*\n?", "", text, flags=re.MULTILINE)
text = f'model = "{MODEL}"\n' + text

config_path.write_text(text, encoding="utf-8")
print(f"  [OK] モデル設定完了（{MODEL}）")

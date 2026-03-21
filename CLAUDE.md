# クリニック書類自動生成システム

## プロジェクト概要

カルテ情報をLLMで構造化JSON に変換し、そのJSONから書類を自動生成するCLIツール。
LLMはローカル（Ollama等）とクラウドAPI（OpenAI互換）を `config.json` で切り替え可能。

## アーキテクチャ

```
カルテテキスト → [Codex CLI / LLM (Python)] → JSON → [書類生成] → .docx / .xls / .xlsx
```

- **Python層**: LLM呼び出し、プロンプト管理、CLI制御 (`generate.py`, `lib/llm_client.py`)
- **LLMプロバイダー**: Codex CLI（デフォルト）、OpenAI互換API、ローカルLLM（Ollama等）
- **Node.js層**: JSON→docxのレンダリング（紹介状・診断書）(`lib/templates/*.js`)
- **Python層（テンプレート書き込み）**: 公式Excelテンプレートにデータを書き込み（自立支援・障害年金）(`lib/templates/*.py`)
  - 自立支援: xlrd + xlutils で .xls テンプレートに書き込み
  - 障害年金: openpyxl で .xlsx テンプレートに書き込み
- 各層はsubprocessで連携。中間データはJSONで受け渡し。

## ディレクトリ構成

```
clinic_docs/
│
│  ★ ユーザーが触る部分 ★
├── config.json              # LLM設定 + クリニック情報（※要編集、gitignore対象）
├── input/                   # カルテ入力フォルダ
│   ├── referral/            #   紹介状バッチ入力用
│   ├── jiritsu_shien/       #   自立支援バッチ入力用
│   └── shougai_nenkin/      #   障害年金バッチ入力用
├── output/                  # 生成ファイル出力先（日付別に自動整理）
│
│  ★ 実行ファイル ★
├── generate.py              # メインCLI（エントリーポイント）
├── run_batch.py             # 汎用バッチ生成スクリプト
├── 紹介状バッチ生成.bat / 自立支援書類作成.bat / 障害年金書類作成.bat
│
│  ★ 内部モジュール ★
├── lib/
│   ├── llm_client.py        #   LLM抽象化レイヤー
│   ├── prompts/             #   LLM用プロンプトテンプレート
│   ├── schemas/             #   JSONスキーマ（各書類の構造定義）
│   └── templates/           #   書類生成スクリプト
│       ├── referral.js          #   紹介状（Node.js/docx-js → .docx）
│       ├── diagnosis_certificate.js  #   診断書（Node.js/docx-js → .docx）
│       ├── jiritsu_shien.py     #   自立支援（Python/xlwings → 公式.xlsに書き込み）
│       ├── jiritsu_shien_word.py  #  自立支援Word版（Python/python-docx → テンプレート.docxに書き込み）
│       └── shougai_nenkin.py    #   障害年金（Python/openpyxl → 公式.xlsxに書き込み）
│
├── sample/                  # 公式テンプレート + テスト用サンプルJSON
│   ├── 自立支援/2jiritsu.xls
│   └── 障害年金/04 (1).xlsx
└── config.example.json      # 設定テンプレート（git追跡）
```

## 対応書類

| キー             | 書類名                           | 出力形式 | 用途                           |
|------------------|----------------------------------|----------|--------------------------------|
| `referral`       | 紹介状（診療情報提供書）         | .docx    | 他院への患者紹介               |
| `diagnosis`      | 診断書                           | .docx    | 保険会社・職場への提出         |
| `jiritsu_shien`  | 自立支援医療診断書（精神通院医療用） | .xls   | 精神通院医療の公費負担申請（公式テンプレートに書き込み） |
| `jiritsu_shien_word` | 自立支援医療診断書 Word版        | .docx  | 自立支援（Word形式で安定出力）     |
| `shougai_nenkin` | 障害年金診断書（精神の障害用・様式第120号の4） | .xlsx | 障害年金の申請（公式テンプレートに書き込み） |

## 使い方

```bash
# カルテテキストから自動生成
python generate.py referral       "カルテのテキスト..."
python generate.py jiritsu_shien  --file karte.txt --save-json

# JSONから直接生成（LLMスキップ）
python generate.py referral --json data.json

# LLMプロバイダー切り替え
python generate.py --provider api referral "..."
```

| オプション         | 短縮  | 説明                                    |
|--------------------|-------|-----------------------------------------|
| `--file`           | `-f`  | カルテテキストファイルのパス             |
| `--json`           | `-j`  | JSONファイルから直接生成（LLMスキップ）  |
| `--output`         | `-o`  | 出力ファイルパス                         |
| `--provider`       | `-p`  | LLMプロバイダー (`codex` / `api` / `local`) |
| `--config`         | `-c`  | 設定ファイルパス                         |
| `--save-json`      |       | 中間JSONファイルも保存                   |

## 技術スタック

- **Python 3.10+**: CLI、LLMクライアント、テンプレート書き込み（xlrd/xlutils, openpyxl, xlwt）
- **Node.js**: docx-js によるWord文書生成（紹介状・診断書）
- **xlrd + xlutils + xlwt**: 自立支援医療診断書の公式.xlsテンプレートへの書き込み
- **openpyxl**: 障害年金診断書の公式.xlsxテンプレートへの書き込み
- **LLM**: Codex CLI（デフォルト）、OpenAI互換API（OpenAI, Ollama, LM Studio 等）

## テンプレート書き込みの注意点

### 自立支援（jiritsu_shien.py）
- xlutils.copy はフォント情報をリセットするため、`xlwt.easyxf` で ＭＳ 明朝 を明示指定
- 症状チェック: テキスト内の項目番号の前に○を挿入する方式
- 条件詳細セルのみ ＭＳ Ｐゴシック 11pt

### 障害年金（shougai_nenkin.py）
- チェックボックス: データバリデーション `"レ,　"` → "レ" で checked
- 結合セルが775箇所あり、必ず結合範囲の左上セルに書き込むこと
- 日付は和暦変換必須（`_era_year()` を使用）
- 治療歴の年数字: U列/AG列に書き込む（Y列/AK列は「年」ラベルなので上書き禁止）
- ICD-10コード: AA26に書き込む（O26はラベルセル）
- 元号セル: ドロップダウンで "昭和","平成","令和" のいずれかを設定
- 患者情報（氏名・住所・生年月日）は空欄のまま（事務が後記入）

## 注意事項

- LLMの出力は必ず医師が確認・修正してから使用すること
- `config.json` にAPIキーを直接書かないこと（環境変数を使用）
- ICD-10コードはLLMが推定するため、正確性を要確認
- LLMモデルは config.json の api.model で設定（現在: chatgpt-5.4）

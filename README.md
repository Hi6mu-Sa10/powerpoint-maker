# PowerPoint Maker MCP Server

Amazon Bedrock (Claude Sonnet) と python-pptx を組み合わせた、Cline専用のPowerPoint自動生成MCPサーバーです。

> **Clineユーザーへ**: このリポジトリには `.clinerules` ファイルが含まれています。
> Clineがこのディレクトリを開いた状態で「セットアップして」と言うだけで、
> Clineが自動的に依存関係のインストール・環境設定・MCPサーバー登録を行います。

---

## 特徴

- **自然言語でPPTX生成**: 「AI活用の提案資料を8枚で」と言うだけで完成
- **Amazon Bedrock (VPCエンドポイント経由)**: Claude Sonnet がスライド構成を動的に生成
- **既存PPTXスタイル適用**: 参照PPTXを渡すとそのデザインを自動抽出・適用
- **テンプレート不要**: LLMが毎回最適な構成を生成するためJSONテンプレート不要
- **stdio方式**: ポート不要・セキュアなローカル通信

---

## ディレクトリ構成

```
powerpoint-maker/
├── mcp-server/
│   ├── server.py              # MCPサーバー本体（stdio）
│   ├── slide_generator.py     # Bedrock Claude連携でスライド構成動的生成
│   ├── style_extractor.py     # 既存PPTXからスタイル抽出
│   └── pptx_builder.py        # python-pptxでPPTXファイル生成
├── output/                    # 生成されたPPTXの出力先
├── .env                       # AWS設定（gitignore対象）
├── .env.example               # 設定サンプル
├── .gitignore
├── requirements.txt
└── README.md
```

---

## セットアップ

### 1. 仮想環境（venv）の作成と依存ライブラリのインストール

```bash
python -m venv .venv
.venv\Scripts\pip install -r requirements.txt
```

> **なぜvenvを使うか？** システムのPython環境を汚染せず、ライブラリのバージョン競合を防ぐためです。

### 2. 環境変数の設定

`.env.example` をコピーして `.env` を作成し、設定値を入力してください。

```bash
copy .env.example .env
```

```.env
AWS_REGION=ap-northeast-1
AWS_PROFILE=<AWSプロファイル名>
BEDROCK_ENDPOINT_URL=https://vpce-xxxxxxxxxx.bedrock-runtime.ap-northeast-1.vpce.amazonaws.com
BEDROCK_MODEL_ID=jp.anthropic.claude-sonnet-4-6
OUTPUT_DIR=output
```

### 3. MCPサーバーの登録（初回のみ）

`cline_mcp_settings.json` に以下を追加（既に設定済みの場合は不要）:

> **重要: `command` には venv内の `python.exe` の絶対パスを指定してください**

```json
"powerpoint-maker": {
  "autoApprove": [
    "generate_pptx",
    "extract_pptx_style",
    "list_slide_types",
    "suggest_slide_structure",
    "build_pptx_from_data"
  ],
  "disabled": false,
  "timeout": 120,
  "type": "stdio",
  "command": "[リポジトリの絶対パス]\\.venv\\Scripts\\python.exe",
  "args": ["[リポジトリの絶対パス]\\mcp-server\\server.py"],
  "env": {
    "AWS_REGION": "ap-northeast-1",
    "AWS_PROFILE": "[AWS SSOプロファイル]",
    "BEDROCK_ENDPOINT_URL": "[BedrockのVPCエンドポイント]",
    "BEDROCK_MODEL_ID": "jp.anthropic.claude-sonnet-4-6",
    "OUTPUT_DIR": "[リポジトリの絶対パス]\\output",
    "PYTHONIOENCODING": "utf-8"
  }
}
```

---

## 使い方（Clineへの指示例）

### 基本的な使い方

```
「AI活用による業務効率化」についての提案資料を10枚で作成してください
```

### 詳細オプション付き

```
以下の条件でPowerPointを作成してください:
- テーマ: クラウド移行プロジェクト提案
- 枚数: 12枚
- 対象: 経営層
- スタイル: フォーマル
- 追加指示: コスト削減効果とROIを強調する
```

### 既存PPTXのスタイルを適用

```
「新製品ローンチ計画」のPPTXを作成してください。
デザインはこのファイルに合わせてください: C:\Users\xxx\template.pptx
```

---

## MCPツール一覧

| ツール名 | 説明 |
|---|---|
| `generate_pptx` | テーマ・条件を指定してPPTXを自動生成（Bedrock連携） |
| `extract_pptx_style` | 既存PPTXからスタイル情報（カラー・フォント）を抽出 |
| `build_pptx_from_data` | スライドJSONデータを直接渡してPPTXを生成 |
| `suggest_slide_structure` | テーマに最適なスライド構成をBedrockが提案 |
| `list_slide_types` | 利用可能なスライドタイプ一覧を表示 |

---

## 利用可能なスライドタイプ

| タイプ | 説明 | 主なフィールド |
|---|---|---|
| `title` | 表紙 | title, subtitle, meta |
| `agenda` | 目次 | title, items[{number, text}] |
| `section` | セクション区切り | title, section_number |
| `content` | 箇条書きコンテンツ | title, bullets[{text, sub}], note |
| `two_column` | 2カラム比較 | title, left{title, items}, right{title, items} |
| `timeline` | ロードマップ | title, milestones[{phase, period, description}] |
| `closing` | クロージング | title, message, contact |

---

## pptx_builder の単体テスト

```bash
python mcp-server/pptx_builder.py
# → output/sample.pptx が生成される
```

## slide_generator の単体テスト（Bedrock接続必要）

```bash
python mcp-server/slide_generator.py "クラウド移行提案" 8
```

---

## アーキテクチャ

```
[Cline (VSCode)] ←stdio→ [MCPサーバー (server.py)]
                                    ↓
                    ┌───────────────────────────────┐
                    │  slide_generator.py            │
                    │  ↓ Amazon Bedrock (VPC経由)    │
                    │  Claude Sonnet がJSONを生成     │
                    └───────────────────────────────┘
                                    ↓
                    ┌───────────────────────────────┐
                    │  style_extractor.py (任意)     │
                    │  参照PPTXからスタイル抽出       │
                    └───────────────────────────────┘
                                    ↓
                    ┌───────────────────────────────┐
                    │  pptx_builder.py               │
                    │  python-pptxでPPTX生成         │
                    └───────────────────────────────┘
                                    ↓
                              output/*.pptx
```

---

## セキュリティ

- MCPサーバーはstdio方式のため**ポートを開放しない**
- Bedrockアクセスは**VPCエンドポイント経由**（インターネット非経由）
- APIキー等の機密情報は `.env` で管理（`.gitignore` で除外済み）

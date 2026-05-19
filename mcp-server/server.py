"""
server.py
PowerPoint生成MCPサーバー（stdio方式）
Clineから直接呼び出せるMCPツールを提供する
"""
from __future__ import annotations

import json
import os
import sys
from datetime import datetime
from pathlib import Path
from typing import Any

# プロジェクトルートをパスに追加
_HERE = Path(__file__).parent
_ROOT = _HERE.parent
sys.path.insert(0, str(_HERE))

from dotenv import load_dotenv

load_dotenv(str(_ROOT / ".env"))

from mcp.server import Server
from mcp.server.stdio import stdio_server
from mcp.types import (
    CallToolResult,
    TextContent,
    Tool,
)

from pptx_builder import SLIDE_RENDERERS, build_pptx
from slide_generator import generate_slides, refine_slide, suggest_slide_types
from style_extractor import extract_style_to_json

# ─────────────────────────────────────────────
# サーバー初期化
# ─────────────────────────────────────────────

app = Server("powerpoint-maker")

OUTPUT_DIR = Path(os.environ.get("OUTPUT_DIR", str(_ROOT / "output")))
OUTPUT_DIR.mkdir(parents=True, exist_ok=True)


# ─────────────────────────────────────────────
# ツール定義
# ─────────────────────────────────────────────

@app.list_tools()
async def list_tools() -> list[Tool]:
    return [
        Tool(
            name="generate_pptx",
            description=(
                "テーマと条件を指定してPowerPointプレゼンテーションを自動生成します。"
                "Amazon Bedrock (Claude Sonnet) がスライド構成を動的に作成し、"
                "python-pptxでPPTXファイルを出力します。"
                "参照PPTXのパスを指定するとそのデザインスタイルを適用します。"
            ),
            inputSchema={
                "type": "object",
                "properties": {
                    "topic": {
                        "type": "string",
                        "description": "プレゼンテーションのテーマ・内容（例: 'AI活用による業務効率化提案'）",
                    },
                    "num_slides": {
                        "type": "integer",
                        "description": "スライド枚数（デフォルト: 8）",
                        "default": 8,
                        "minimum": 3,
                        "maximum": 20,
                    },
                    "audience": {
                        "type": "string",
                        "description": "対象オーディエンス（例: '経営層', '技術者', '営業チーム'）",
                        "default": "",
                    },
                    "style_hint": {
                        "type": "string",
                        "description": "スタイルのヒント（例: 'フォーマル', 'カジュアル', '簡潔に'）",
                        "default": "",
                    },
                    "extra_instructions": {
                        "type": "string",
                        "description": "追加指示（例: '競合比較を含める', 'KPIを強調する'）",
                        "default": "",
                    },
                    "reference_pptx": {
                        "type": "string",
                        "description": "デザイン参照用PPTXファイルのパス（省略可）",
                        "default": "",
                    },
                    "output_filename": {
                        "type": "string",
                        "description": "出力ファイル名（.pptx拡張子含む。省略時は自動命名）",
                        "default": "",
                    },
                },
                "required": ["topic"],
            },
        ),
        Tool(
            name="extract_pptx_style",
            description=(
                "既存のPPTXファイルからデザインスタイル情報（カラーパレット・フォント・"
                "レイアウト等）を抽出してJSON形式で返します。"
                "抽出したスタイルはgenerate_pptxのデザイン参照として活用できます。"
            ),
            inputSchema={
                "type": "object",
                "properties": {
                    "pptx_path": {
                        "type": "string",
                        "description": "スタイルを抽出するPPTXファイルのパス",
                    },
                    "save_json": {
                        "type": "boolean",
                        "description": "抽出結果をJSONファイルとして保存するか（デフォルト: false）",
                        "default": False,
                    },
                },
                "required": ["pptx_path"],
            },
        ),
        Tool(
            name="list_slide_types",
            description=(
                "利用可能なスライドタイプと各タイプの説明一覧を返します。"
                "スライド構成を手動でカスタマイズする際に参照してください。"
            ),
            inputSchema={
                "type": "object",
                "properties": {},
            },
        ),
        Tool(
            name="suggest_slide_structure",
            description=(
                "テーマを入力すると、そのプレゼンテーションに最適なスライド構成を"
                "Amazon Bedrock (Claude Sonnet) が提案します。"
            ),
            inputSchema={
                "type": "object",
                "properties": {
                    "topic": {
                        "type": "string",
                        "description": "プレゼンテーションのテーマ",
                    },
                },
                "required": ["topic"],
            },
        ),
        Tool(
            name="build_pptx_from_data",
            description=(
                "スライド構成データ（JSON）を直接渡してPPTXを生成します。"
                "generate_pptxで生成したスライドデータを編集してから"
                "PPTXを作り直したい場合に使います。"
            ),
            inputSchema={
                "type": "object",
                "properties": {
                    "slides_json": {
                        "type": "string",
                        "description": "スライドデータのJSON文字列（slides配列またはスライドオブジェクトの配列）",
                    },
                    "reference_pptx": {
                        "type": "string",
                        "description": "デザイン参照用PPTXファイルのパス（省略可）",
                        "default": "",
                    },
                    "output_filename": {
                        "type": "string",
                        "description": "出力ファイル名（省略時は自動命名）",
                        "default": "",
                    },
                },
                "required": ["slides_json"],
            },
        ),
    ]


# ─────────────────────────────────────────────
# ツール実装
# ─────────────────────────────────────────────

@app.call_tool()
async def call_tool(name: str, arguments: dict[str, Any]) -> list[TextContent]:

    try:
        if name == "generate_pptx":
            return await _handle_generate_pptx(arguments)
        elif name == "extract_pptx_style":
            return await _handle_extract_style(arguments)
        elif name == "list_slide_types":
            return await _handle_list_slide_types(arguments)
        elif name == "suggest_slide_structure":
            return await _handle_suggest_structure(arguments)
        elif name == "build_pptx_from_data":
            return await _handle_build_from_data(arguments)
        else:
            return [TextContent(type="text", text=f"❌ 未知のツール: {name}")]

    except Exception as e:
        import traceback
        tb = traceback.format_exc()
        return [TextContent(
            type="text",
            text=f"❌ エラーが発生しました: {type(e).__name__}: {e}\n\n詳細:\n{tb}"
        )]


async def _handle_generate_pptx(args: dict[str, Any]) -> list[TextContent]:
    """generate_pptxハンドラー"""
    topic = args["topic"]
    num_slides = args.get("num_slides", 8)
    audience = args.get("audience", "")
    style_hint = args.get("style_hint", "")
    extra_instructions = args.get("extra_instructions", "")
    reference_pptx = args.get("reference_pptx", "")
    output_filename = args.get("output_filename", "")

    # 出力ファイル名の決定
    if not output_filename:
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        safe_topic = "".join(c for c in topic[:20] if c.isalnum() or c in "ー_")
        output_filename = f"{timestamp}_{safe_topic}.pptx"
    elif not output_filename.endswith(".pptx"):
        output_filename += ".pptx"

    output_path = str(OUTPUT_DIR / output_filename)

    # ステップ1: スタイル抽出（参照PPTXがある場合）
    style = None
    style_info = ""
    if reference_pptx and Path(reference_pptx).exists():
        try:
            style_json_str = extract_style_to_json(reference_pptx)
            style = json.loads(style_json_str)
            style_info = f"\n📐 参照PPTXのスタイルを適用: {Path(reference_pptx).name}"
        except Exception as e:
            style_info = f"\n⚠️  スタイル抽出をスキップ（エラー: {e}）"

    # ステップ2: スライド構成をBedrockで生成
    slides = generate_slides(
        topic=topic,
        num_slides=num_slides,
        style_hint=style_hint,
        audience=audience,
        extra_instructions=extra_instructions,
    )

    # ステップ3: PPTXを生成
    final_path = build_pptx(slides, output_path, style=style)

    # 結果レポート
    slide_summary = _format_slide_summary(slides)
    result = f"""✅ PowerPointを生成しました！

📁 出力ファイル: {final_path}
📊 スライド枚数: {len(slides)}枚
🎯 テーマ: {topic}{style_info}

## スライド構成
{slide_summary}

## 次のステップ
- ファイルを開いて内容を確認してください
- 修正が必要な場合は `build_pptx_from_data` ツールで再生成できます
- デザインの参照PPTXを指定すると既存スタイルを適用できます"""

    return [TextContent(type="text", text=result)]


async def _handle_extract_style(args: dict[str, Any]) -> list[TextContent]:
    """extract_pptx_styleハンドラー"""
    pptx_path = args["pptx_path"]
    save_json = args.get("save_json", False)

    if not Path(pptx_path).exists():
        return [TextContent(
            type="text",
            text=f"❌ ファイルが見つかりません: {pptx_path}"
        )]

    output_json_path = None
    if save_json:
        stem = Path(pptx_path).stem
        output_json_path = str(OUTPUT_DIR / f"{stem}_style.json")

    style_json = extract_style_to_json(pptx_path, output_json_path)
    style = json.loads(style_json)

    # 見やすいサマリーを生成
    summary_lines = [
        f"✅ スタイル抽出完了: {Path(pptx_path).name}",
        "",
        "## 抽出された情報",
        f"- スライドサイズ: {style.get('slide_width_pt')}pt × {style.get('slide_height_pt')}pt",
        f"- テーマカラー数: {len(style.get('theme_colors', {}))}色",
        f"- 検出フォント数: {len(style.get('fonts', {}).get('detected', []))}種類",
        f"- スライドレイアウト数: {len(style.get('slide_layouts', []))}種類",
    ]

    fonts = style.get("fonts", {})
    if fonts.get("title"):
        summary_lines.append(f"- タイトルフォント: {fonts['title']}")
    if fonts.get("body"):
        summary_lines.append(f"- 本文フォント: {fonts['body']}")

    theme_colors = style.get("theme_colors", {})
    if theme_colors:
        summary_lines.append("\n## テーマカラー")
        for k, v in list(theme_colors.items())[:6]:
            summary_lines.append(f"  - {k}: {v}")

    if save_json and output_json_path:
        summary_lines.append(f"\n📄 JSONを保存: {output_json_path}")

    summary_lines.append("\n## 生スタイルデータ（JSON）")
    summary_lines.append("```json")
    summary_lines.append(json.dumps(
        {k: v for k, v in style.items() if k not in ["sample_text_styles", "slide_layouts"]},
        ensure_ascii=False, indent=2
    ))
    summary_lines.append("```")

    return [TextContent(type="text", text="\n".join(summary_lines))]


async def _handle_list_slide_types(args: dict[str, Any]) -> list[TextContent]:
    """list_slide_typesハンドラー"""
    slide_types = {
        "title": {
            "説明": "表紙スライド。タイトル・サブタイトル・日付・発表者名を配置",
            "フィールド": ["title", "subtitle", "meta"],
        },
        "agenda": {
            "説明": "目次スライド。番号付きアジェンダ一覧を表示",
            "フィールド": ["title", "items[{number, text}]"],
        },
        "section": {
            "説明": "セクション区切りスライド。章の始まりを視覚的に示す",
            "フィールド": ["title", "section_number"],
        },
        "content": {
            "説明": "箇条書きコンテンツスライド。メインポイントと補足説明を表示",
            "フィールド": ["title", "bullets[{text, sub}]", "note"],
        },
        "two_column": {
            "説明": "2カラム比較スライド。左右に項目を並べて比較",
            "フィールド": ["title", "left{title, items}", "right{title, items}"],
        },
        "timeline": {
            "説明": "タイムライン・ロードマップスライド。フェーズと期間を時系列で表示",
            "フィールド": ["title", "milestones[{phase, period, description}]"],
        },
        "closing": {
            "説明": "クロージングスライド。Thank Youメッセージと連絡先を表示",
            "フィールド": ["title", "message", "contact"],
        },
    }

    lines = ["# 利用可能なスライドタイプ\n"]
    for type_name, info in slide_types.items():
        lines.append(f"## `{type_name}`")
        lines.append(f"**説明**: {info['説明']}")
        lines.append(f"**フィールド**: {', '.join(info['フィールド'])}")
        lines.append("")

    return [TextContent(type="text", text="\n".join(lines))]


async def _handle_suggest_structure(args: dict[str, Any]) -> list[TextContent]:
    """suggest_slide_structureハンドラー"""
    topic = args["topic"]
    types = suggest_slide_types(topic)

    lines = [
        f"# 「{topic}」の推奨スライド構成\n",
        f"総スライド数: {len(types)}枚\n",
        "## スライド構成",
    ]
    for i, t in enumerate(types, 1):
        lines.append(f"{i}. `{t}`")

    lines.append("\n## 説明")
    type_desc = {
        "title": "表紙",
        "agenda": "目次",
        "section": "セクション区切り",
        "content": "箇条書きコンテンツ",
        "two_column": "2カラム比較",
        "timeline": "タイムライン・ロードマップ",
        "closing": "クロージング",
    }
    for t in types:
        lines.append(f"- **{t}**: {type_desc.get(t, t)}")

    lines.append(
        f"\n💡 この構成でPPTXを生成するには:\n"
        f"```\ngenerate_pptx ツールで topic=\"{topic}\" と指定してください\n```"
    )

    return [TextContent(type="text", text="\n".join(lines))]


async def _handle_build_from_data(args: dict[str, Any]) -> list[TextContent]:
    """build_pptx_from_dataハンドラー"""
    slides_json_str = args["slides_json"]
    reference_pptx = args.get("reference_pptx", "")
    output_filename = args.get("output_filename", "")

    # JSONパース
    try:
        data = json.loads(slides_json_str)
    except json.JSONDecodeError as e:
        return [TextContent(type="text", text=f"❌ JSON解析エラー: {e}")]

    # slides配列の取り出し
    if isinstance(data, dict) and "slides" in data:
        slides = data["slides"]
    elif isinstance(data, list):
        slides = data
    else:
        return [TextContent(type="text", text="❌ スライドデータの形式が不正です。配列またはslidesキーを持つオブジェクトを渡してください。")]

    # 出力ファイル名
    if not output_filename:
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        output_filename = f"{timestamp}_custom.pptx"
    elif not output_filename.endswith(".pptx"):
        output_filename += ".pptx"

    output_path = str(OUTPUT_DIR / output_filename)

    # スタイル抽出
    style = None
    style_info = ""
    if reference_pptx and Path(reference_pptx).exists():
        try:
            style_json_str = extract_style_to_json(reference_pptx)
            style = json.loads(style_json_str)
            style_info = f"\n📐 参照スタイル: {Path(reference_pptx).name}"
        except Exception as e:
            style_info = f"\n⚠️ スタイル抽出スキップ: {e}"

    # PPTX生成
    final_path = build_pptx(slides, output_path, style=style)

    slide_summary = _format_slide_summary(slides)
    result = f"""✅ PowerPointを生成しました！

📁 出力ファイル: {final_path}
📊 スライド枚数: {len(slides)}枚{style_info}

## スライド構成
{slide_summary}"""

    return [TextContent(type="text", text=result)]


def _format_slide_summary(slides: list[dict[str, Any]]) -> str:
    """スライド構成のサマリーテキストを生成"""
    lines = []
    for i, slide in enumerate(slides, 1):
        t = slide.get("type", "?")
        title = slide.get("title", "（タイトルなし）")
        type_icon = {
            "title": "🎯",
            "agenda": "📋",
            "section": "📌",
            "content": "📝",
            "two_column": "⬛",
            "timeline": "📅",
            "closing": "🙏",
        }.get(t, "▪️")
        lines.append(f"{i:2d}. {type_icon} [{t}] {title}")
    return "\n".join(lines)


# ─────────────────────────────────────────────
# エントリーポイント
# ─────────────────────────────────────────────

async def main() -> None:
    async with stdio_server() as (read_stream, write_stream):
        await app.run(
            read_stream,
            write_stream,
            app.create_initialization_options(),
        )


if __name__ == "__main__":
    import asyncio
    asyncio.run(main())

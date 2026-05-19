"""
slide_generator.py
Amazon Bedrock (Claude Sonnet) を使って、自然言語からスライド構成データを動的生成するモジュール
"""
from __future__ import annotations

import json
import os
from pathlib import Path
import re
from typing import Any

import boto3
from dotenv import load_dotenv

# .envを読み込み
_ENV_PATH = Path(__file__).parent.parent / ".env"
load_dotenv(str(_ENV_PATH))

# ─────────────────────────────────────────────
# Bedrock クライアント初期化
# ─────────────────────────────────────────────

def _get_bedrock_client() -> Any:
    """Amazon Bedrockクライアントを返す"""
    region = os.environ.get("AWS_REGION", "ap-northeast-1")
    endpoint_url = os.environ.get("BEDROCK_ENDPOINT_URL")
    profile = os.environ.get("AWS_PROFILE")

    session = boto3.Session(
        profile_name=profile,
        region_name=region,
    )

    kwargs: dict[str, Any] = {
        "service_name": "bedrock-runtime",
        "region_name": region,
    }
    if endpoint_url:
        kwargs["endpoint_url"] = endpoint_url

    return session.client(**kwargs)


# ─────────────────────────────────────────────
# プロンプト定義
# ─────────────────────────────────────────────

SYSTEM_PROMPT = """あなたはプロフェッショナルなプレゼンテーション設計の専門家です。
ユーザーの要望を聞いて、PowerPointスライドの構成データをJSONで生成してください。

## 出力フォーマット

以下のJSONスキーマに厳密に従ってください。余分なテキストは出力しないでください。

```json
{
  "slides": [
    // タイトルスライド（必須・最初）
    {
      "type": "title",
      "title": "プレゼンタイトル",
      "subtitle": "サブタイトルや概要",
      "meta": "日付　発表者名　所属"
    },
    // 目次スライド
    {
      "type": "agenda",
      "title": "目次",
      "items": [
        {"number": 1, "text": "アジェンダ項目1"},
        {"number": 2, "text": "アジェンダ項目2"}
      ]
    },
    // セクション区切り
    {
      "type": "section",
      "section_number": "01",
      "title": "セクションタイトル"
    },
    // 箇条書きコンテンツ
    {
      "type": "content",
      "title": "スライドタイトル",
      "bullets": [
        {"text": "メインポイント", "sub": "補足説明（省略可）"},
        {"text": "メインポイント2"}
      ],
      "note": "補足メモ（省略可）"
    },
    // 2カラム比較
    {
      "type": "two_column",
      "title": "スライドタイトル",
      "left": {
        "title": "左カラムタイトル",
        "items": ["項目1", "項目2"]
      },
      "right": {
        "title": "右カラムタイトル",
        "items": ["項目1", "項目2"]
      }
    },
    // タイムライン・ロードマップ
    {
      "type": "timeline",
      "title": "ロードマップ",
      "milestones": [
        {"phase": "Phase 1", "period": "2025年Q3", "description": "内容説明"},
        {"phase": "Phase 2", "period": "2025年Q4", "description": "内容説明"}
      ]
    },
    // クロージング（必須・最後）
    {
      "type": "closing",
      "title": "Thank You",
      "message": "ご清聴ありがとうございました",
      "contact": "連絡先情報"
    }
  ]
}
```

## ルール
- 必ずJSONのみを出力（前後の説明文は不要）
- タイトルスライドとクロージングスライドは必須
- スライド枚数はユーザー指定に従う（指定なければ7〜10枚）
- 日本語でリアルで具体的なコンテンツを生成
- 箇条書きは1スライドあたり最大5項目
- タイムラインは最大5マイルストーン
- 目次項目は実際のコンテンツスライドと対応させる
"""


def _call_bedrock(prompt: str, system: str = SYSTEM_PROMPT) -> str:
    """Bedrock Claude APIを呼び出してテキストを返す"""
    client = _get_bedrock_client()
    model_id = os.environ.get("BEDROCK_MODEL_ID", "jp.anthropic.claude-sonnet-4-6")

    # システムプロンプトにキャッシュを適用（トークン削減）
    system_with_cache = [
        {
            "type": "text",
            "text": system,
            "cache_control": {"type": "ephemeral"},
        }
    ]

    body = {
        "anthropic_version": "bedrock-2023-05-31",
        "max_tokens": 8192,
        "system": system_with_cache,
        "messages": [
            {"role": "user", "content": prompt}
        ],
        "temperature": 0.7,
    }

    response = client.invoke_model(
        modelId=model_id,
        body=json.dumps(body),
        contentType="application/json",
        accept="application/json",
    )

    result = json.loads(response["body"].read())
    return result["content"][0]["text"]


def _extract_json(text: str) -> dict[str, Any]:
    """レスポンステキストからJSONを抽出してパース"""
    # コードブロック内のJSONを探す
    patterns = [
        r"```json\s*([\s\S]+?)\s*```",
        r"```\s*([\s\S]+?)\s*```",
        r"(\{[\s\S]+\})",
    ]
    for pattern in patterns:
        match = re.search(pattern, text)
        if match:
            try:
                return json.loads(match.group(1))
            except json.JSONDecodeError:
                continue

    # そのままパース
    try:
        return json.loads(text.strip())
    except json.JSONDecodeError as e:
        raise ValueError(f"JSONの解析に失敗しました: {e}\n\nレスポンス:\n{text[:500]}") from e


def generate_slides(
    topic: str,
    num_slides: int = 8,
    style_hint: str = "",
    audience: str = "",
    extra_instructions: str = "",
) -> list[dict[str, Any]]:
    """
    トピックからスライド構成データを生成

    Args:
        topic: プレゼンテーションのテーマ・内容
        num_slides: 希望スライド枚数（デフォルト8）
        style_hint: スタイルのヒント（例: "フォーマル", "カジュアル"）
        audience: 対象オーディエンス（例: "経営層", "技術者"）
        extra_instructions: 追加指示

    Returns:
        スライドデータのリスト
    """
    parts = [f"テーマ: {topic}"]
    parts.append(f"スライド枚数: {num_slides}枚（タイトル・目次・クロージング含む）")

    if audience:
        parts.append(f"対象オーディエンス: {audience}")
    if style_hint:
        parts.append(f"スタイル: {style_hint}")
    if extra_instructions:
        parts.append(f"追加指示: {extra_instructions}")

    parts.append("\n上記の条件でプレゼンテーションのスライド構成JSONを生成してください。")

    prompt = "\n".join(parts)
    raw = _call_bedrock(prompt)
    data = _extract_json(raw)
    slides = data.get("slides", [])

    if not slides:
        raise ValueError("スライドデータが空です。Bedrockからの応答を確認してください。")

    return slides


def refine_slide(
    slide_data: dict[str, Any],
    instruction: str,
) -> dict[str, Any]:
    """
    既存のスライドデータを指示に基づいて改善

    Args:
        slide_data: 改善対象のスライドデータ
        instruction: 改善指示

    Returns:
        改善されたスライドデータ
    """
    system = """あなたはプレゼンテーション編集の専門家です。
既存のスライドデータを指示に基づいて改善し、同じJSONフォーマットで返してください。
必ずJSONのみを出力してください（前後の説明文は不要）。"""

    prompt = f"""以下のスライドデータを改善してください。

## 現在のスライドデータ
```json
{json.dumps(slide_data, ensure_ascii=False, indent=2)}
```

## 改善指示
{instruction}

改善後のスライドデータをJSONで返してください。"""

    raw = _call_bedrock(prompt, system=system)
    data = _extract_json(raw)

    # 単一スライドの場合は直接返す
    if "slides" in data:
        slides = data["slides"]
        return slides[0] if slides else slide_data
    return data


def suggest_slide_types(topic: str) -> list[str]:
    """
    トピックに最適なスライド構成タイプを提案

    Args:
        topic: プレゼンテーションのテーマ

    Returns:
        推奨スライドタイプのリスト
    """
    system = """あなたはプレゼンテーション設計の専門家です。
トピックに最適なスライド構成を提案してください。
必ずJSON配列のみを返してください。"""

    prompt = f"""テーマ「{topic}」のプレゼンテーションに最適なスライド構成を提案してください。

利用可能なスライドタイプ:
- title: タイトルスライド
- agenda: 目次
- section: セクション区切り
- content: 箇条書きコンテンツ
- two_column: 2カラム比較
- timeline: タイムライン・ロードマップ
- closing: クロージング

8〜10枚のスライド構成をJSON配列で返してください。
例: ["title", "agenda", "section", "content", "two_column", "closing"]"""

    raw = _call_bedrock(prompt, system=system)
    try:
        match = re.search(r"\[[\s\S]+\]", raw)
        if match:
            return json.loads(match.group(0))
    except Exception:
        pass
    return ["title", "agenda", "section", "content", "content", "two_column", "timeline", "closing"]


if __name__ == "__main__":
    import sys

    topic = sys.argv[1] if len(sys.argv) > 1 else "AI活用による業務効率化"
    num = int(sys.argv[2]) if len(sys.argv) > 2 else 8

    print(f"🔄 スライド生成中: '{topic}' ({num}枚)...")
    try:
        slides = generate_slides(topic, num_slides=num)
        print(f"✅ {len(slides)}枚のスライドを生成しました")
        print(json.dumps(slides, ensure_ascii=False, indent=2))
    except Exception as e:
        print(f"❌ エラー: {e}", file=sys.stderr)
        sys.exit(1)

"""
Gemini API Analyzer
Gemini APIを使用して画像からセクションを検出・分析するモジュール
"""

import google.generativeai as genai
from PIL import Image
from typing import List, Dict, Optional, Tuple
import json
import os
import base64
from io import BytesIO


class GeminiAnalyzer:
    """Gemini APIを使用して画像を分析するクラス"""

    def __init__(self, api_key: str, model_name: str = "gemini-1.5-flash"):
        """
        初期化

        Args:
            api_key: Gemini APIキー
            model_name: 使用するモデル名（デフォルト: gemini-1.5-flash）
        """
        self.api_key = api_key
        self.model_name = model_name

        # Gemini APIの設定
        genai.configure(api_key=api_key)
        self.model = genai.GenerativeModel(model_name)

    def detect_sections(self, image_path: str) -> List[Dict]:
        """
        画像からセクションを検出

        Args:
            image_path: 分析する画像のパス

        Returns:
            検出されたセクションのリスト
            [
                {
                    "section_id": 1,
                    "title": "セクションタイトル",
                    "description": "セクションの説明",
                    "bounds": {"top": 0, "left": 0, "bottom": 100, "right": 100}
                },
                ...
            ]
        """
        try:
            # 画像を読み込み
            img = Image.open(image_path)

            # プロンプトを作成
            prompt = """
この画像はExcelシートを描画したものです。
この画像を分析して、論理的なセクション（意味のあるまとまり）に区切ってください。

各セクションについて、以下の情報をJSON形式で返してください：
- section_id: セクションのID（1から始まる連番）
- title: セクションのタイトルまたは主題
- description: セクションの簡単な説明
- row_range: 行の範囲（例: "1-5"）
- content_type: コンテンツの種類（"table", "header", "summary", "data"など）

JSON形式で返してください。例：
{
    "sections": [
        {
            "section_id": 1,
            "title": "ヘッダー情報",
            "description": "タイトルと基本情報",
            "row_range": "1-3",
            "content_type": "header"
        },
        {
            "section_id": 2,
            "title": "売上データ",
            "description": "月別の売上データテーブル",
            "row_range": "4-15",
            "content_type": "table"
        }
    ]
}
"""

            # Gemini APIで分析
            response = self.model.generate_content([prompt, img])

            # レスポンスをパース
            try:
                # JSONとして解析
                response_text = response.text.strip()

                # マークダウンのコードブロックを除去
                if response_text.startswith("```json"):
                    response_text = response_text[7:]
                if response_text.startswith("```"):
                    response_text = response_text[3:]
                if response_text.endswith("```"):
                    response_text = response_text[:-3]

                response_text = response_text.strip()

                result = json.loads(response_text)
                sections = result.get("sections", [])

                return sections

            except json.JSONDecodeError as e:
                print(f"JSON解析エラー: {e}")
                print(f"レスポンス: {response.text}")
                # フォールバック: 単一セクションとして返す
                return [{
                    "section_id": 1,
                    "title": "全体",
                    "description": response.text[:200],
                    "row_range": "全体",
                    "content_type": "unknown"
                }]

        except Exception as e:
            print(f"セクション検出エラー: {e}")
            return []

    def analyze_section(self, image_path: str, section_info: Optional[Dict] = None) -> Dict:
        """
        セクション画像を詳細に分析

        Args:
            image_path: 分析する画像のパス
            section_info: セクション情報（オプション）

        Returns:
            分析結果
            {
                "summary": "要約",
                "details": "詳細な説明",
                "key_points": ["ポイント1", "ポイント2", ...],
                "data_structure": "データ構造の説明",
                "insights": "インサイト"
            }
        """
        try:
            # 画像を読み込み
            img = Image.open(image_path)

            # プロンプトを作成
            if section_info:
                context = f"""
セクション情報:
- タイトル: {section_info.get('title', '不明')}
- 説明: {section_info.get('description', '不明')}
- 行範囲: {section_info.get('row_range', '不明')}
- コンテンツタイプ: {section_info.get('content_type', '不明')}
"""
            else:
                context = "このセクションについて"

            prompt = f"""
{context}

この画像（Excelシートの一部）を詳細に分析してください。

以下の項目についてJSON形式で返してください：
- summary: 全体の要約（50-100文字）
- details: 詳細な説明（200-300文字）
- key_points: 重要なポイントのリスト（3-5個）
- data_structure: データ構造や表の構成の説明
- insights: データから読み取れるインサイトや特徴
- markdown_table: 可能であれば、Markdown形式のテーブル

JSON形式で返してください。例：
{{
    "summary": "2023年度の月別売上データ",
    "details": "1月から12月までの売上データを含むテーブル。各月の売上高、利益率、前年比が記載されている。",
    "key_points": [
        "12月の売上が最も高い",
        "利益率は平均15%",
        "前年比で平均8%の成長"
    ],
    "data_structure": "12行×4列のテーブル。列は月、売上高、利益率、前年比。",
    "insights": "年末にかけて売上が増加する傾向が見られる。",
    "markdown_table": "| 月 | 売上高 | 利益率 | 前年比 |\\n|---|---|---|---|\\n| 1月 | 100万円 | 15% | +5% |"
}}
"""

            # Gemini APIで分析
            response = self.model.generate_content([prompt, img])

            # レスポンスをパース
            try:
                response_text = response.text.strip()

                # マークダウンのコードブロックを除去
                if response_text.startswith("```json"):
                    response_text = response_text[7:]
                if response_text.startswith("```"):
                    response_text = response_text[3:]
                if response_text.endswith("```"):
                    response_text = response_text[:-3]

                response_text = response_text.strip()

                result = json.loads(response_text)
                return result

            except json.JSONDecodeError as e:
                print(f"JSON解析エラー: {e}")
                print(f"レスポンス: {response.text}")
                # フォールバック
                return {
                    "summary": "分析結果（テキスト形式）",
                    "details": response.text,
                    "key_points": [],
                    "data_structure": "不明",
                    "insights": ""
                }

        except Exception as e:
            print(f"セクション分析エラー: {e}")
            return {
                "summary": "エラー",
                "details": f"分析中にエラーが発生しました: {str(e)}",
                "key_points": [],
                "data_structure": "不明",
                "insights": ""
            }

    def analyze_full_sheet(self, image_path: str) -> Dict:
        """
        シート全体を分析（セクション検出 + 各セクション分析）

        Args:
            image_path: 分析する画像のパス

        Returns:
            完全な分析結果
            {
                "sections": [
                    {
                        "section_info": {...},
                        "analysis": {...}
                    },
                    ...
                ],
                "overall_summary": "全体の要約"
            }
        """
        results = {
            "sections": [],
            "overall_summary": ""
        }

        try:
            # ステップ1: セクションを検出
            sections = self.detect_sections(image_path)

            # ステップ2: 各セクションを分析
            # （実際には各セクションを切り出して分析する必要があるが、
            # ここでは画像全体を各セクションのコンテキストで分析）
            for section in sections:
                analysis = self.analyze_section(image_path, section)

                results["sections"].append({
                    "section_info": section,
                    "analysis": analysis
                })

            # ステップ3: 全体の要約を生成
            img = Image.open(image_path)
            prompt = "この画像（Excelシート）全体の内容を100文字程度で要約してください。"
            response = self.model.generate_content([prompt, img])
            results["overall_summary"] = response.text.strip()

        except Exception as e:
            print(f"シート全体分析エラー: {e}")
            results["overall_summary"] = f"エラー: {str(e)}"

        return results

    def generate_markdown_from_analysis(self, analysis_results: Dict, sheet_name: str) -> str:
        """
        分析結果からMarkdownを生成

        Args:
            analysis_results: analyze_full_sheetの結果
            sheet_name: シート名

        Returns:
            Markdown形式のテキスト
        """
        md_lines = []

        # タイトル
        md_lines.append(f"# {sheet_name}")
        md_lines.append("")

        # 全体の要約
        md_lines.append("## 全体の要約")
        md_lines.append(analysis_results.get("overall_summary", ""))
        md_lines.append("")

        # 各セクション
        for idx, section_data in enumerate(analysis_results.get("sections", []), 1):
            section_info = section_data.get("section_info", {})
            analysis = section_data.get("analysis", {})

            # セクションヘッダー
            section_title = section_info.get("title", f"セクション {idx}")
            md_lines.append(f"## {section_title}")
            md_lines.append("")

            # セクション情報
            md_lines.append(f"**行範囲:** {section_info.get('row_range', '不明')}")
            md_lines.append(f"**タイプ:** {section_info.get('content_type', '不明')}")
            md_lines.append("")

            # 要約
            md_lines.append(f"**要約:** {analysis.get('summary', '')}")
            md_lines.append("")

            # 詳細
            md_lines.append("### 詳細")
            md_lines.append(analysis.get('details', ''))
            md_lines.append("")

            # 重要ポイント
            key_points = analysis.get('key_points', [])
            if key_points:
                md_lines.append("### 重要ポイント")
                for point in key_points:
                    md_lines.append(f"- {point}")
                md_lines.append("")

            # データ構造
            if analysis.get('data_structure'):
                md_lines.append("### データ構造")
                md_lines.append(analysis.get('data_structure', ''))
                md_lines.append("")

            # インサイト
            if analysis.get('insights'):
                md_lines.append("### インサイト")
                md_lines.append(analysis.get('insights', ''))
                md_lines.append("")

            # Markdownテーブル（もしあれば）
            if analysis.get('markdown_table'):
                md_lines.append("### データテーブル")
                md_lines.append(analysis.get('markdown_table', ''))
                md_lines.append("")

            md_lines.append("---")
            md_lines.append("")

        return "\n".join(md_lines)

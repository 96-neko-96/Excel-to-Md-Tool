"""
Sheet Parser - シート解析ロジック
"""

from typing import Dict, Any, List
import openpyxl
from .table_parser import TableParser
from .image_parser import ImageParser


class SheetParser:
    """シート解析クラス"""

    def __init__(self, config: Dict[str, Any], gemini_analyzer=None):
        self.config = config
        self.table_parser = TableParser(config)
        self.image_parser = ImageParser(config)
        self.gemini_analyzer = gemini_analyzer  # Phase 3: AI機能用

    def set_gemini_analyzer(self, gemini_analyzer):
        """Phase 3: GeminiAnalyzerを設定"""
        self.gemini_analyzer = gemini_analyzer

    def parse_sheet(self, sheet, sheet_with_values=None) -> Dict[str, Any]:
        """
        シートを解析してMarkdown形式に変換

        Args:
            sheet: openpyxlのWorksheetオブジェクト（数式情報用）
            sheet_with_values: 実数値を含むシート（data_only=True）

        Returns:
            シートデータの辞書
        """
        sheet_data = {
            'name': sheet.title,
            'index': sheet.sheet_properties.sheetId if hasattr(sheet.sheet_properties, 'sheetId') else 0,
            'content': '',
            'cell_range': self._get_used_range(sheet),
            'tables': [],
            'images': [],
            'shapes': [],
            'tables_count': 0,
            'images_count': 0,
            'shapes_count': 0
        }

        try:
            # テーブルの検出と変換
            tables_md, tables_info = self.table_parser.parse_tables(sheet, sheet_with_values)
            sheet_data['tables'] = tables_info
            sheet_data['tables_count'] = len(tables_info)

            # 画像の抽出（設定により）
            images_md = []
            if self.config.get('extract_images', True):
                images_md, images_info = self.image_parser.extract_images(sheet)
                sheet_data['images'] = images_info
                sheet_data['images_count'] = len(images_info)

            # 図形の抽出（常に実行）
            shapes_md, shapes_info = self.image_parser.extract_shapes(sheet)
            sheet_data['shapes'] = shapes_info
            sheet_data['shapes_count'] = len(shapes_info)

            # コンテンツの結合
            content_parts = []

            # テーブルを追加
            if tables_md:
                content_parts.extend(tables_md)

            # 画像を追加
            if images_md:
                content_parts.extend(images_md)

            # 図形を追加
            if shapes_md:
                content_parts.extend(shapes_md)

            # もしテーブルも画像もない場合は、シート全体をテーブルとして扱う
            if not content_parts:
                fallback_md = self._convert_sheet_as_table(sheet, sheet_with_values)
                if fallback_md:
                    content_parts.append(fallback_md)

            sheet_data['content'] = '\n\n'.join(content_parts)

            # Phase 3: AI機能による追加コンテンツ生成
            if self.gemini_analyzer and self.config.get('enable_ai_features'):
                try:
                    ai_sections = []

                    # 表の要約（全テーブルをセクションとしてまとめて追加）
                    if self.config.get('ai_table_summary') and tables_info:
                        table_summaries = []
                        for idx, table_info in enumerate(tables_info):
                            if 'markdown' in table_info:
                                summary = self.gemini_analyzer.generate_table_summary(
                                    table_info['markdown']
                                )
                                table_summaries.append({
                                    'table_index': idx,
                                    'summary': summary,
                                    'table_name': table_info.get('name', f'Table {idx + 1}')
                                })

                        if table_summaries:
                            sheet_data['table_summaries'] = table_summaries
                            # AI要約セクションを生成
                            summary_section = self._format_table_summaries_section(table_summaries)
                            ai_sections.append(summary_section)

                    # 画像の説明（全画像をセクションとしてまとめて追加）
                    if self.config.get('ai_image_description') and images_info:
                        image_descriptions = []
                        for idx, image_info in enumerate(images_info):
                            if 'path' in image_info:
                                description = self.gemini_analyzer.generate_image_description(
                                    image_info['path']
                                )
                                image_descriptions.append({
                                    'image_index': idx,
                                    'description': description,
                                    'image_name': image_info.get('name', f'Image {idx + 1}')
                                })

                        if image_descriptions:
                            sheet_data['image_descriptions'] = image_descriptions
                            # AI説明セクションを生成
                            description_section = self._format_image_descriptions_section(image_descriptions)
                            ai_sections.append(description_section)

                    # QA生成（シート全体の最後に追加）
                    if self.config.get('ai_generate_qa') and sheet_data['content']:
                        qa_list = self.gemini_analyzer.generate_qa_for_sheet(
                            sheet_data['content'],
                            sheet.title
                        )
                        if qa_list:
                            qa_md = self._format_qa_section(qa_list)
                            ai_sections.append(qa_md)
                            sheet_data['qa_list'] = qa_list

                    # すべてのAI生成コンテンツを本文に追加
                    if ai_sections:
                        sheet_data['content'] += '\n\n' + '\n\n'.join(ai_sections)

                except Exception as e:
                    print(f"AI機能エラー（シート: {sheet.title}）: {e}")
                    # AI機能のエラーは致命的ではないので続行

        except Exception as e:
            import traceback
            error_msg = f"シート '{sheet.title}' の解析エラー: {str(e)}"
            print(error_msg)

            # デバッグモードの場合は詳細なエラー情報を出力
            if self.config.get('verbose_logging', False):
                traceback.print_exc()

            # エラーメッセージをコンテンツに追加
            sheet_data['content'] = f"⚠️ このシートの解析中にエラーが発生しました: {str(e)}"

        return sheet_data

    def _format_table_summaries_section(self, table_summaries: List[Dict[str, Any]]) -> str:
        """Phase 3: 表の要約セクションをMarkdown形式でフォーマット"""
        lines = [
            "\n---\n",
            "## 🤖 AI生成: 表の要約\n",
            "> **注意**: 以下の内容はAIによって自動生成されたものです。\n"
        ]

        for item in table_summaries:
            table_name = item.get('table_name', f"Table {item['table_index'] + 1}")
            summary = item.get('summary', '')
            lines.append(f"### 📊 {table_name}\n")
            lines.append(f"{summary}\n")

        return '\n'.join(lines)

    def _format_image_descriptions_section(self, image_descriptions: List[Dict[str, Any]]) -> str:
        """Phase 3: 画像説明セクションをMarkdown形式でフォーマット"""
        lines = [
            "\n---\n",
            "## 🤖 AI生成: 画像の説明\n",
            "> **注意**: 以下の内容はAIによって自動生成されたものです。\n"
        ]

        for item in image_descriptions:
            image_name = item.get('image_name', f"Image {item['image_index'] + 1}")
            description = item.get('description', '')
            lines.append(f"### 🖼️ {image_name}\n")
            lines.append(f"{description}\n")

        return '\n'.join(lines)

    def _format_qa_section(self, qa_list: List[Dict[str, str]]) -> str:
        """Phase 3: QAセクションをMarkdown形式でフォーマット"""
        lines = [
            "\n---\n",
            "## 🤖 AI生成: よくある質問\n",
            "> **注意**: 以下の内容はAIによって自動生成されたものです。\n"
        ]

        for idx, qa in enumerate(qa_list, 1):
            lines.append(f"### ❓ Q{idx}: {qa.get('question', '')}\n")
            lines.append(f"**A:** {qa.get('answer', '')}\n")

        return '\n'.join(lines)

    def _get_used_range(self, sheet) -> str:
        """使用されているセル範囲を取得"""
        if sheet.dimensions:
            return sheet.dimensions
        return "A1:A1"

    def _convert_sheet_as_table(self, sheet, sheet_with_values=None) -> str:
        """シート全体を1つのテーブルとして変換（フォールバック）"""
        # 使用されている範囲を取得
        if not sheet.dimensions or sheet.dimensions == "A1:A1":
            return ""

        # シート全体を1つの大きなテーブルとして解析
        return self.table_parser.convert_range_to_markdown(sheet, sheet.dimensions, sheet_with_values)

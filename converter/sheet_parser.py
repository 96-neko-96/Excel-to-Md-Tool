"""
Sheet Parser - シート解析ロジック
"""

from typing import Dict, Any, List
import openpyxl
from .table_parser import TableParser
from .image_parser import ImageParser


class SheetParser:
    """シート解析クラス"""

    def __init__(self, config: Dict[str, Any]):
        self.config = config
        self.table_parser = TableParser(config)
        self.image_parser = ImageParser(config)

    def parse_sheet(self, sheet) -> Dict[str, Any]:
        """
        シートを解析してMarkdown形式に変換

        Args:
            sheet: openpyxlのWorksheetオブジェクト

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
            'tables_count': 0,
            'images_count': 0
        }

        # テーブルの検出と変換
        tables_md, tables_info = self.table_parser.parse_tables(sheet)
        sheet_data['tables'] = tables_info
        sheet_data['tables_count'] = len(tables_info)

        # 画像の抽出（設定により）
        images_md = []
        if self.config.get('extract_images', True):
            images_md, images_info = self.image_parser.extract_images(sheet)
            sheet_data['images'] = images_info
            sheet_data['images_count'] = len(images_info)

        # コンテンツの結合
        content_parts = []

        # テーブルを追加
        if tables_md:
            content_parts.extend(tables_md)

        # 画像を追加
        if images_md:
            content_parts.extend(images_md)

        # もしテーブルも画像もない場合は、シート全体をテーブルとして扱う
        if not content_parts:
            fallback_md = self._convert_sheet_as_table(sheet)
            if fallback_md:
                content_parts.append(fallback_md)

        sheet_data['content'] = '\n\n'.join(content_parts)

        return sheet_data

    def _get_used_range(self, sheet) -> str:
        """使用されているセル範囲を取得"""
        if sheet.dimensions:
            return sheet.dimensions
        return "A1:A1"

    def _convert_sheet_as_table(self, sheet) -> str:
        """シート全体を1つのテーブルとして変換（フォールバック）"""
        # 使用されている範囲を取得
        if not sheet.dimensions or sheet.dimensions == "A1:A1":
            return ""

        # シート全体を1つの大きなテーブルとして解析
        return self.table_parser.convert_range_to_markdown(sheet, sheet.dimensions)

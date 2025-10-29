"""
Excel to Markdown Converter Core
"""

import os
import openpyxl
from typing import Dict, List, Any
from datetime import datetime

from .sheet_parser import SheetParser
from .markdown_generator import MarkdownGenerator
from .metadata_generator import MetadataGenerator
from .image_parser import ImageParser


class ExcelToMarkdownConverter:
    """Excel to Markdown 変換エンジン"""

    def __init__(self, **config):
        """
        Args:
            chunk_size: RAG用チャンクサイズ
            create_toc: 目次生成の有効化
            extract_images: 画像抽出の有効化
            generate_summary: 表要約の有効化
        """
        self.config = {
            'chunk_size': config.get('chunk_size', 800),
            'create_toc': config.get('create_toc', True),
            'extract_images': config.get('extract_images', True),
            'generate_summary': config.get('generate_summary', False),
        }

        self.workbook = None
        self.sheets_data = []
        self.cross_references = []

        # パーサーの初期化
        self.sheet_parser = SheetParser(self.config)
        self.markdown_generator = MarkdownGenerator(self.config)
        self.metadata_generator = MetadataGenerator(self.config)
        self.image_parser = ImageParser(self.config)

    def convert(self, input_path: str, output_path: str) -> Dict[str, Any]:
        """
        Excelファイルを変換

        Args:
            input_path: 入力Excelファイルパス
            output_path: 出力Markdownファイルパス

        Returns:
            変換結果の統計情報
        """
        # 1. Excel読み込み
        self.workbook = self._load_excel(input_path)

        # 2. 各シートを変換
        for sheet in self.workbook.worksheets:
            # 非表示シートはスキップ（設定による）
            if sheet.sheet_state == 'hidden' and not self.config.get('include_hidden', False):
                continue

            sheet_data = self._convert_sheet(sheet)
            self.sheets_data.append(sheet_data)

        # 3. シート間参照解析
        self.cross_references = self._analyze_references()

        # 4. 統合Markdown生成
        md_content = self._merge_sheets()

        # 5. ファイル出力
        self._write_output(output_path, md_content)

        # 6. メタデータ生成
        metadata = self._generate_metadata(input_path, output_path)

        return {
            'sheets_count': len(self.sheets_data),
            'tables_count': sum(s['tables_count'] for s in self.sheets_data),
            'images_count': sum(s['images_count'] for s in self.sheets_data),
            'estimated_chunks': self._estimate_chunks(md_content),
            'metadata': metadata,
            'output_file': output_path
        }

    def _load_excel(self, path: str) -> openpyxl.Workbook:
        """Excelファイルを読み込む"""
        try:
            # data_onlyをFalseにして数式も読み込む
            workbook = openpyxl.load_workbook(path, data_only=False)
            return workbook
        except Exception as e:
            raise ValueError(f"Excelファイルの読み込みに失敗しました: {str(e)}")

    def _convert_sheet(self, sheet) -> Dict[str, Any]:
        """シートを変換"""
        return self.sheet_parser.parse_sheet(sheet)

    def _analyze_references(self) -> List[Dict[str, Any]]:
        """シート間参照を解析"""
        references = []

        if not self.workbook:
            return references

        for sheet in self.workbook.worksheets:
            for row in sheet.iter_rows():
                for cell in row:
                    if cell.data_type == 'f':  # 数式セル
                        formula = str(cell.value)

                        # 他シート参照の検出
                        if '!' in formula:
                            ref_info = self._parse_cross_sheet_reference(formula, sheet.title, cell.coordinate)
                            if ref_info:
                                references.append(ref_info)

        return references

    def _parse_cross_sheet_reference(self, formula: str, from_sheet: str, from_cell: str) -> Dict[str, Any]:
        """数式から他シート参照を解析"""
        import re

        # 例: "=SUM(Sheet1!A1:A10)" または "='Sheet Name'!A1"
        pattern = r"(['\"]?)([^'\"!]+)\1!([A-Z]+[0-9]+(?::[A-Z]+[0-9]+)?)"
        matches = re.findall(pattern, formula)

        if matches:
            # 最初の参照のみを取得
            _, sheet_name, cell_range = matches[0]
            return {
                'from_sheet': from_sheet,
                'from_cell': from_cell,
                'to_sheet': sheet_name,
                'to_cell': cell_range,
                'formula': formula
            }

        return None

    def _merge_sheets(self) -> str:
        """複数シートを統合"""
        return self.markdown_generator.merge_sheets(
            self.sheets_data,
            self.cross_references,
            self.workbook
        )

    def _write_output(self, path: str, content: str):
        """Markdownファイルを出力"""
        try:
            output_dir = os.path.dirname(path)
            if output_dir and not os.path.exists(output_dir):
                os.makedirs(output_dir)

            with open(path, 'w', encoding='utf-8') as f:
                f.write(content)
        except Exception as e:
            raise IOError(f"ファイルの書き込みに失敗しました: {str(e)}")

    def _generate_metadata(self, input_path: str, output_path: str) -> Dict[str, Any]:
        """メタデータを生成"""
        return self.metadata_generator.generate(
            self.workbook,
            self.sheets_data,
            self.cross_references,
            input_path,
            output_path
        )

    def _estimate_chunks(self, content: str) -> int:
        """推奨チャンク数を推定"""
        # 簡易的な推定（1文字 ≒ 0.3トークン程度）
        estimated_tokens = len(content) * 0.3
        chunk_size = self.config['chunk_size']

        return max(1, int(estimated_tokens // chunk_size) + 1)

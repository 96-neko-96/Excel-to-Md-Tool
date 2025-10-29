"""
Metadata Generator - メタデータ生成ロジック
"""

import os
from typing import Dict, List, Any
from datetime import datetime
import re
from collections import Counter


class MetadataGenerator:
    """メタデータ生成クラス"""

    def __init__(self, config: Dict[str, Any]):
        self.config = config

    def generate(self, workbook, sheets_data: List[Dict[str, Any]],
                cross_references: List[Dict[str, Any]],
                input_path: str, output_path: str) -> Dict[str, Any]:
        """
        RAG用メタデータをJSON形式で生成

        Args:
            workbook: openpyxlのWorkbookオブジェクト
            sheets_data: シートデータのリスト
            cross_references: シート間参照のリスト
            input_path: 入力ファイルパス
            output_path: 出力ファイルパス

        Returns:
            メタデータ辞書
        """
        metadata = {
            'source_file': os.path.basename(input_path),
            'source_path': input_path,
            'converted_at': datetime.now().isoformat(),
            'output_file': os.path.basename(output_path),
            'output_path': output_path,
            'total_sheets': len(sheets_data),
            'sheets': [],
            'cross_references': [],
            'statistics': {
                'total_tables': 0,
                'total_images': 0,
                'total_size_kb': 0
            },
            'rag_optimization': {
                'chunk_size': self.config.get('chunk_size', 800),
                'estimated_chunks': 0
            }
        }

        # 各シートのメタデータ
        for sheet_data in sheets_data:
            sheet_meta = {
                'name': sheet_data['name'],
                'index': sheet_data['index'],
                'cell_range': sheet_data['cell_range'],
                'tables_count': sheet_data['tables_count'],
                'images_count': sheet_data['images_count'],
                'section_in_md': f"#{self._create_anchor(sheet_data['name'])}",
                'keywords': self._extract_keywords(sheet_data['content'])
            }
            metadata['sheets'].append(sheet_meta)

            metadata['statistics']['total_tables'] += sheet_meta['tables_count']
            metadata['statistics']['total_images'] += sheet_meta['images_count']

        # シート間参照
        for ref in cross_references:
            metadata['cross_references'].append({
                'from': f"{ref['from_sheet']}!{ref['from_cell']}",
                'to': f"{ref['to_sheet']}!{ref['to_cell']}",
                'type': self._detect_reference_type(ref['formula'])
            })

        # ファイルサイズ
        if os.path.exists(output_path):
            file_size = os.path.getsize(output_path)
            metadata['statistics']['total_size_kb'] = round(file_size / 1024, 2)

        # チャンク数推定
        if os.path.exists(output_path):
            with open(output_path, 'r', encoding='utf-8') as f:
                content = f.read()
            total_tokens = self._estimate_tokens(content)
            chunk_size = self.config.get('chunk_size', 800)
            metadata['rag_optimization']['estimated_chunks'] = max(1, int(total_tokens // chunk_size) + 1)

        return metadata

    def _create_anchor(self, sheet_name: str) -> str:
        """アンカー名を生成"""
        anchor = sheet_name.replace(' ', '-').replace('_', '-').lower()
        anchor = ''.join(c for c in anchor if c.isalnum() or c == '-')
        return anchor

    def _extract_keywords(self, content: str) -> List[str]:
        """コンテンツからキーワードを抽出"""
        if not content or not self.config.get('extract_keywords', True):
            return []

        # 簡易的なキーワード抽出
        # 日本語と英語の単語を抽出
        words = re.findall(r'[a-zA-Z]+|[ぁ-んァ-ヶー一-龠]+', content)

        # 長さ2文字以上の単語のみ
        words = [w for w in words if len(w) >= 2]

        # 頻度の高い単語を抽出（上位10個）
        word_counts = Counter(words)
        keywords = [word for word, count in word_counts.most_common(10)]

        return keywords

    def _detect_reference_type(self, formula: str) -> str:
        """数式のタイプを検出"""
        formula_upper = formula.upper()

        if 'SUM' in formula_upper:
            return 'sum'
        elif 'AVERAGE' in formula_upper or 'AVG' in formula_upper:
            return 'average'
        elif 'COUNT' in formula_upper:
            return 'count'
        elif 'VLOOKUP' in formula_upper or 'HLOOKUP' in formula_upper:
            return 'lookup'
        elif 'IF' in formula_upper:
            return 'conditional'
        else:
            return 'reference'

    def _estimate_tokens(self, content: str) -> int:
        """トークン数を推定"""
        # 簡易的な推定（1文字 ≒ 0.3トークン）
        # 日本語は1文字 ≒ 0.5トークン、英語は1単語 ≒ 1.3トークン程度
        return int(len(content) * 0.3)

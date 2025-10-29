"""
Markdown Generator - Markdown生成ロジック
"""

from typing import List, Dict, Any
from datetime import datetime
import os


class MarkdownGenerator:
    """Markdown生成クラス"""

    def __init__(self, config: Dict[str, Any]):
        self.config = config

    def merge_sheets(self, sheets_data: List[Dict[str, Any]],
                     cross_references: List[Dict[str, Any]],
                     workbook) -> str:
        """
        複数シートを1つのMarkdownファイルに統合

        Args:
            sheets_data: シートデータのリスト
            cross_references: シート間参照のリスト
            workbook: openpyxlのWorkbookオブジェクト

        Returns:
            統合されたMarkdownコンテンツ
        """
        merged = []

        # ヘッダー情報
        header = self._generate_header(workbook, sheets_data)
        merged.append(header)

        # 目次生成
        if self.config.get('create_toc', True):
            toc = self._generate_toc(sheets_data)
            merged.append(toc)
            merged.append("\n---\n")

        # 各シートのコンテンツ
        for idx, sheet_data in enumerate(sheets_data, 1):
            # シート区切り
            if idx > 1:
                merged.append("\n---\n")

            # シート見出し
            sheet_anchor = self._create_anchor(sheet_data['name'])
            merged.append(f"\n<a name=\"{sheet_anchor}\"></a>")
            merged.append(f"# {idx}. {sheet_data['name']}\n")

            # シートのコンテンツ
            if sheet_data['content']:
                merged.append(sheet_data['content'])
            else:
                merged.append("*このシートには表示可能なコンテンツがありません*")

            # シート間参照情報を追加
            related_refs = self._find_related_references(
                sheet_data['name'],
                cross_references
            )
            if related_refs:
                ref_text = self._generate_reference_links(related_refs)
                merged.append(f"\n{ref_text}\n")

        return '\n'.join(merged)

    def _generate_header(self, workbook, sheets_data: List[Dict[str, Any]]) -> str:
        """ファイルヘッダーを生成"""
        header_parts = []

        # タイトル
        title = getattr(workbook.properties, 'title', None) or 'Excel Document'
        header_parts.append(f"# {title}\n")

        # ファイル情報
        header_parts.append("**ファイル情報**")
        header_parts.append(f"- シート数: {len(sheets_data)}")
        header_parts.append(f"- 変換日時: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")

        # 作成者情報（あれば）
        if hasattr(workbook.properties, 'creator') and workbook.properties.creator:
            header_parts.append(f"- 作成者: {workbook.properties.creator}")

        header_parts.append("")  # 空行

        return '\n'.join(header_parts)

    def _generate_toc(self, sheets_data: List[Dict[str, Any]]) -> str:
        """目次を生成"""
        toc = ["## 目次\n"]

        for idx, sheet_data in enumerate(sheets_data, 1):
            sheet_anchor = self._create_anchor(sheet_data['name'])

            # シート名とテーブル・画像数の情報
            info_parts = []
            if sheet_data['tables_count'] > 0:
                info_parts.append(f"{sheet_data['tables_count']}表")
            if sheet_data['images_count'] > 0:
                info_parts.append(f"{sheet_data['images_count']}画像")

            info_str = f" ({', '.join(info_parts)})" if info_parts else ""

            toc.append(f"{idx}. [{sheet_data['name']}](#{sheet_anchor}){info_str}")

        return '\n'.join(toc)

    def _create_anchor(self, sheet_name: str) -> str:
        """アンカー名を生成"""
        # スペースをハイフンに、特殊文字を削除
        anchor = sheet_name.replace(' ', '-').replace('_', '-').lower()
        # 英数字とハイフン以外を削除
        anchor = ''.join(c for c in anchor if c.isalnum() or c == '-')
        return anchor

    def _find_related_references(self, sheet_name: str,
                                 cross_references: List[Dict[str, Any]]) -> List[Dict[str, Any]]:
        """指定されたシートに関連する参照を検索"""
        related = []
        for ref in cross_references:
            if ref['from_sheet'] == sheet_name:
                related.append({
                    'direction': 'to',
                    'sheet': ref['to_sheet'],
                    'cell': ref['to_cell']
                })
            elif ref['to_sheet'] == sheet_name:
                related.append({
                    'direction': 'from',
                    'sheet': ref['from_sheet'],
                    'cell': ref['from_cell']
                })
        return related

    def _generate_reference_links(self, references: List[Dict[str, Any]]) -> str:
        """シート間参照リンクを生成"""
        if not references:
            return ""

        links = ["\n> **関連シート:**"]

        for ref in references:
            anchor = self._create_anchor(ref['sheet'])
            direction = "参照先" if ref['direction'] == 'to' else "参照元"
            links.append(f"> - [{ref['sheet']}](#{anchor}) ({direction}: セル {ref['cell']})")

        return '\n'.join(links)

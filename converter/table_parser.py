"""
Table Parser - 表検出・変換ロジック
"""

from typing import List, Tuple, Dict, Any
import pandas as pd
import openpyxl
from openpyxl.utils import get_column_letter


class TableParser:
    """テーブル解析・変換クラス"""

    def __init__(self, config: Dict[str, Any]):
        self.config = config

    def parse_tables(self, sheet) -> Tuple[List[str], List[Dict[str, Any]]]:
        """
        シート内のテーブルを検出して変換

        Args:
            sheet: openpyxlのWorksheetオブジェクト

        Returns:
            (Markdown形式のテーブルリスト, テーブル情報のリスト)
        """
        tables_md = []
        tables_info = []

        # Excelの正式なテーブルオブジェクトがあればそれを使用
        if hasattr(sheet, 'tables') and sheet.tables:
            for table_name, table_range in sheet.tables.items():
                md_table = self.convert_range_to_markdown(sheet, table_range)
                if md_table:
                    tables_md.append(md_table)
                    tables_info.append({
                        'name': table_name,
                        'range': table_range,
                        'type': 'excel_table'
                    })
        else:
            # テーブルオブジェクトがない場合は、使用範囲全体を1つのテーブルとして扱う
            if sheet.dimensions and sheet.dimensions != "A1:A1":
                md_table = self.convert_range_to_markdown(sheet, sheet.dimensions)
                if md_table:
                    tables_md.append(md_table)
                    tables_info.append({
                        'name': 'data',
                        'range': sheet.dimensions,
                        'type': 'auto_detected'
                    })

        return tables_md, tables_info

    def convert_range_to_markdown(self, sheet, cell_range: str) -> str:
        """
        セル範囲をMarkdownテーブルに変換

        Args:
            sheet: openpyxlのWorksheetオブジェクト
            cell_range: セル範囲（例: "A1:D10"）

        Returns:
            Markdown形式のテーブル
        """
        try:
            # セル範囲からデータを取得
            data = []
            for row in sheet[cell_range]:
                row_data = []
                for cell in row:
                    # セルの値を取得（数式の場合は計算結果を使用）
                    value = cell.value
                    if value is None:
                        value = ""
                    row_data.append(str(value))
                data.append(row_data)

            if not data:
                return ""

            # pandasのDataFrameに変換
            df = pd.DataFrame(data)

            # 空の行と列を削除
            df = df.replace('', pd.NA).dropna(how='all', axis=0).dropna(how='all', axis=1)

            if df.empty:
                return ""

            # ヘッダーの検出
            if self.config.get('detect_header', True) and len(df) > 0:
                # 最初の行をヘッダーとして使用
                headers = df.iloc[0].tolist()
                df = df.iloc[1:]
                df.columns = headers
            else:
                # ヘッダーなし（列番号を使用）
                df.columns = [f"Column {i+1}" for i in range(len(df.columns))]

            # Markdown形式に変換
            md_table = df.to_markdown(index=False, tablefmt='github')

            # 要約の生成（設定により）
            if self.config.get('generate_summary', False):
                summary = self._generate_table_summary(df)
                if summary:
                    md_table = f"{summary}\n\n{md_table}"

            return md_table

        except Exception as e:
            print(f"テーブル変換エラー: {str(e)}")
            return ""

    def _generate_table_summary(self, df: pd.DataFrame) -> str:
        """テーブルの要約を生成"""
        summary_parts = []

        # 基本情報
        summary_parts.append(f"データ行数: {len(df)}行")

        # 数値列の情報
        numeric_cols = df.select_dtypes(include='number').columns.tolist()
        if numeric_cols:
            summary_parts.append(f"数値列: {', '.join(numeric_cols)}")

        if summary_parts:
            return "【テーブル要約】 " + "、".join(summary_parts)

        return ""

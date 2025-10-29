"""
Batch Processor - バッチ処理機能
"""

import os
from typing import List, Dict, Any, Callable
from pathlib import Path
import sys

# 親ディレクトリをパスに追加
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from converter import ExcelToMarkdownConverter


class BatchProcessor:
    """バッチ処理クラス"""

    def __init__(self, **default_config):
        """
        Args:
            **default_config: デフォルトの変換設定
        """
        self.default_config = default_config
        self.results = []

    def process_directory(self, input_dir: str, output_dir: str,
                         recursive: bool = False,
                         progress_callback: Callable = None) -> List[Dict[str, Any]]:
        """
        ディレクトリ内のExcelファイルを一括変換

        Args:
            input_dir: 入力ディレクトリ
            output_dir: 出力ディレクトリ
            recursive: サブディレクトリも処理するか
            progress_callback: 進捗コールバック関数 (current, total, filename)

        Returns:
            変換結果のリスト
        """
        # 出力ディレクトリの作成
        os.makedirs(output_dir, exist_ok=True)

        # Excelファイルの検索
        excel_files = self._find_excel_files(input_dir, recursive)

        total_files = len(excel_files)
        self.results = []

        for idx, excel_file in enumerate(excel_files, 1):
            if progress_callback:
                progress_callback(idx, total_files, os.path.basename(excel_file))

            # 出力ファイル名の生成
            relative_path = os.path.relpath(excel_file, input_dir)
            output_filename = os.path.splitext(relative_path)[0] + '.md'
            output_path = os.path.join(output_dir, output_filename)

            # 出力ディレクトリの作成（サブディレクトリ対応）
            os.makedirs(os.path.dirname(output_path), exist_ok=True)

            # 変換実行
            try:
                result = self._process_file(excel_file, output_path)
                result['status'] = 'success'
            except Exception as e:
                result = {
                    'input_file': excel_file,
                    'output_file': output_path,
                    'status': 'error',
                    'error_message': str(e)
                }

            self.results.append(result)

        return self.results

    def process_files(self, input_files: List[str], output_dir: str,
                     progress_callback: Callable = None) -> List[Dict[str, Any]]:
        """
        指定されたファイルリストを一括変換

        Args:
            input_files: 入力ファイルのリスト
            output_dir: 出力ディレクトリ
            progress_callback: 進捗コールバック関数

        Returns:
            変換結果のリスト
        """
        os.makedirs(output_dir, exist_ok=True)

        total_files = len(input_files)
        self.results = []

        for idx, input_file in enumerate(input_files, 1):
            if progress_callback:
                progress_callback(idx, total_files, os.path.basename(input_file))

            # 出力ファイル名の生成
            output_filename = os.path.splitext(os.path.basename(input_file))[0] + '.md'
            output_path = os.path.join(output_dir, output_filename)

            # 変換実行
            try:
                result = self._process_file(input_file, output_path)
                result['status'] = 'success'
            except Exception as e:
                result = {
                    'input_file': input_file,
                    'output_file': output_path,
                    'status': 'error',
                    'error_message': str(e)
                }

            self.results.append(result)

        return self.results

    def _process_file(self, input_file: str, output_file: str) -> Dict[str, Any]:
        """単一ファイルを処理"""
        converter = ExcelToMarkdownConverter(**self.default_config)
        result = converter.convert(input_file, output_file)
        result['input_file'] = input_file
        return result

    def _find_excel_files(self, directory: str, recursive: bool = False) -> List[str]:
        """ディレクトリ内のExcelファイルを検索"""
        excel_extensions = ['.xlsx', '.xls']
        excel_files = []

        if recursive:
            # 再帰的に検索
            for root, dirs, files in os.walk(directory):
                for file in files:
                    if any(file.endswith(ext) for ext in excel_extensions):
                        excel_files.append(os.path.join(root, file))
        else:
            # カレントディレクトリのみ
            for file in os.listdir(directory):
                if any(file.endswith(ext) for ext in excel_extensions):
                    excel_files.append(os.path.join(directory, file))

        return sorted(excel_files)

    def get_summary(self) -> Dict[str, Any]:
        """処理結果のサマリーを取得"""
        if not self.results:
            return {
                'total': 0,
                'success': 0,
                'failed': 0,
                'total_sheets': 0,
                'total_tables': 0,
                'total_images': 0
            }

        success_count = sum(1 for r in self.results if r.get('status') == 'success')
        failed_count = len(self.results) - success_count

        total_sheets = sum(r.get('sheets_count', 0) for r in self.results if r.get('status') == 'success')
        total_tables = sum(r.get('tables_count', 0) for r in self.results if r.get('status') == 'success')
        total_images = sum(r.get('images_count', 0) for r in self.results if r.get('status') == 'success')

        return {
            'total': len(self.results),
            'success': success_count,
            'failed': failed_count,
            'total_sheets': total_sheets,
            'total_tables': total_tables,
            'total_images': total_images,
            'results': self.results
        }

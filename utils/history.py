"""
Conversion History Manager - 変換履歴管理
"""

import json
import os
from typing import List, Dict, Any
from datetime import datetime


class HistoryManager:
    """変換履歴管理クラス"""

    def __init__(self, history_file: str = "conversion_history.json"):
        self.history_file = history_file
        self.history = self._load_history()

    def _load_history(self) -> List[Dict[str, Any]]:
        """履歴ファイルを読み込む"""
        if os.path.exists(self.history_file):
            try:
                with open(self.history_file, 'r', encoding='utf-8') as f:
                    return json.load(f)
            except Exception as e:
                print(f"履歴の読み込みエラー: {e}")
                return []
        else:
            return []

    def save_history(self):
        """履歴をファイルに保存"""
        try:
            with open(self.history_file, 'w', encoding='utf-8') as f:
                json.dump(self.history, f, ensure_ascii=False, indent=2)
        except Exception as e:
            raise IOError(f"履歴の保存に失敗しました: {e}")

    def add_record(self, record: Dict[str, Any]):
        """履歴レコードを追加"""
        record['timestamp'] = datetime.now().isoformat()
        record['id'] = len(self.history) + 1

        # 履歴を追加（最新を先頭に）
        self.history.insert(0, record)

        # 履歴の上限（最大100件）
        if len(self.history) > 100:
            self.history = self.history[:100]

        self.save_history()

    def get_recent(self, limit: int = 10) -> List[Dict[str, Any]]:
        """最近の履歴を取得"""
        return self.history[:limit]

    def get_all(self) -> List[Dict[str, Any]]:
        """すべての履歴を取得"""
        return self.history

    def search(self, keyword: str) -> List[Dict[str, Any]]:
        """キーワードで履歴を検索"""
        results = []
        for record in self.history:
            # ファイル名で検索
            if keyword.lower() in record.get('input_file', '').lower():
                results.append(record)
            # 出力ファイルで検索
            elif keyword.lower() in record.get('output_file', '').lower():
                results.append(record)

        return results

    def delete_record(self, record_id: int):
        """履歴レコードを削除"""
        self.history = [r for r in self.history if r.get('id') != record_id]
        self.save_history()

    def clear_all(self):
        """すべての履歴をクリア"""
        self.history = []
        self.save_history()

    def get_statistics(self) -> Dict[str, Any]:
        """履歴の統計情報を取得"""
        if not self.history:
            return {
                'total_conversions': 0,
                'total_sheets': 0,
                'total_tables': 0,
                'total_images': 0,
                'average_sheets': 0,
                'average_tables': 0,
                'average_images': 0
            }

        total_sheets = sum(r.get('sheets_count', 0) for r in self.history)
        total_tables = sum(r.get('tables_count', 0) for r in self.history)
        total_images = sum(r.get('images_count', 0) for r in self.history)

        count = len(self.history)

        return {
            'total_conversions': count,
            'total_sheets': total_sheets,
            'total_tables': total_tables,
            'total_images': total_images,
            'average_sheets': round(total_sheets / count, 2) if count > 0 else 0,
            'average_tables': round(total_tables / count, 2) if count > 0 else 0,
            'average_images': round(total_images / count, 2) if count > 0 else 0
        }

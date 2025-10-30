"""
Preset Manager - プリセット設定管理
"""

import json
import os
from typing import Dict, List, Any
from datetime import datetime


class PresetManager:
    """プリセット設定管理クラス"""

    def __init__(self, presets_file: str = "presets.json", config_file: str = "config.json"):
        self.presets_file = presets_file
        self.config_file = config_file
        self.presets = self._load_presets()
        self.config = self._load_config()

    def _load_presets(self) -> Dict[str, Any]:
        """プリセットファイルを読み込む"""
        if os.path.exists(self.presets_file):
            try:
                with open(self.presets_file, 'r', encoding='utf-8') as f:
                    return json.load(f)
            except Exception as e:
                print(f"プリセットの読み込みエラー: {e}")
                return self._get_default_presets()
        else:
            return self._get_default_presets()

    def _get_default_presets(self) -> Dict[str, Any]:
        """デフォルトプリセットを取得"""
        return {
            "デフォルト": {
                "chunk_size": 800,
                "create_toc": True,
                "extract_images": True,
                "generate_summary": False,
                "show_formulas": True,
                # Phase 3: AI機能設定（各機能は個別に制御）
                "ai_table_summary": False,
                "ai_image_description": False,
                "ai_generate_qa": False,
                "description": "標準的な変換設定"
            },
            "RAG最適化": {
                "chunk_size": 600,
                "create_toc": True,
                "extract_images": False,
                "generate_summary": True,
                "show_formulas": False,
                # Phase 3: AI機能設定（各機能は個別に制御）
                "ai_table_summary": False,
                "ai_image_description": False,
                "ai_generate_qa": False,
                "description": "RAGシステム向けの最適化設定（画像なし、要約あり）"
            },
            "完全変換": {
                "chunk_size": 1200,
                "create_toc": True,
                "extract_images": True,
                "generate_summary": True,
                "show_formulas": True,
                # Phase 3: AI機能設定（各機能は個別に制御）
                "ai_table_summary": False,
                "ai_image_description": False,
                "ai_generate_qa": False,
                "description": "すべての情報を含む完全変換"
            },
            "軽量版": {
                "chunk_size": 1000,
                "create_toc": False,
                "extract_images": False,
                "generate_summary": False,
                "show_formulas": False,
                # Phase 3: AI機能設定（各機能は個別に制御）
                "ai_table_summary": False,
                "ai_image_description": False,
                "ai_generate_qa": False,
                "description": "最小限の情報のみ（表データのみ）"
            }
        }

    def save_presets(self):
        """プリセットをファイルに保存"""
        try:
            with open(self.presets_file, 'w', encoding='utf-8') as f:
                json.dump(self.presets, f, ensure_ascii=False, indent=2)
        except Exception as e:
            raise IOError(f"プリセットの保存に失敗しました: {e}")

    def get_preset(self, name: str) -> Dict[str, Any]:
        """指定されたプリセットを取得"""
        if name in self.presets:
            return self.presets[name].copy()
        else:
            raise ValueError(f"プリセット '{name}' が見つかりません")

    def get_preset_names(self) -> List[str]:
        """利用可能なプリセット名のリストを取得"""
        return list(self.presets.keys())

    def add_preset(self, name: str, settings: Dict[str, Any], description: str = ""):
        """新しいプリセットを追加"""
        self.presets[name] = settings.copy()
        self.presets[name]["description"] = description
        self.presets[name]["created_at"] = datetime.now().isoformat()
        self.save_presets()

    def delete_preset(self, name: str):
        """プリセットを削除"""
        if name in self.presets:
            del self.presets[name]
            self.save_presets()
        else:
            raise ValueError(f"プリセット '{name}' が見つかりません")

    def update_preset(self, name: str, settings: Dict[str, Any], description: str = None):
        """既存のプリセットを更新"""
        if name not in self.presets:
            raise ValueError(f"プリセット '{name}' が見つかりません")

        self.presets[name].update(settings)
        if description is not None:
            self.presets[name]["description"] = description
        self.presets[name]["updated_at"] = datetime.now().isoformat()
        self.save_presets()

    # Phase 3: グローバル設定管理
    def _load_config(self) -> Dict[str, Any]:
        """グローバル設定を読み込む"""
        if os.path.exists(self.config_file):
            try:
                with open(self.config_file, 'r', encoding='utf-8') as f:
                    return json.load(f)
            except Exception as e:
                print(f"設定の読み込みエラー: {e}")
                return self._get_default_config()
        else:
            return self._get_default_config()

    def _get_default_config(self) -> Dict[str, Any]:
        """デフォルトのグローバル設定を取得"""
        return {
            "gemini_api_key": "",
            "gemini_model": "gemini-2.5-flash-lite",
            "gemini_dpi": 150
        }

    def save_config(self):
        """グローバル設定をファイルに保存"""
        try:
            with open(self.config_file, 'w', encoding='utf-8') as f:
                json.dump(self.config, f, ensure_ascii=False, indent=2)
        except Exception as e:
            raise IOError(f"設定の保存に失敗しました: {e}")

    def get_config(self, key: str = None) -> Any:
        """グローバル設定を取得"""
        if key is None:
            return self.config.copy()
        return self.config.get(key)

    def update_config(self, key: str, value: Any):
        """グローバル設定を更新"""
        self.config[key] = value
        self.save_config()

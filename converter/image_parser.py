"""
Image Parser - 画像・グラフ抽出ロジック
"""

import os
from typing import List, Tuple, Dict, Any
from PIL import Image
import io


class ImageParser:
    """画像・グラフ抽出クラス"""

    def __init__(self, config: Dict[str, Any]):
        self.config = config
        self.image_counter = 0

    def extract_images(self, sheet) -> Tuple[List[str], List[Dict[str, Any]]]:
        """
        シートから画像・グラフを抽出

        Args:
            sheet: openpyxlのWorksheetオブジェクト

        Returns:
            (Markdown形式の画像参照リスト, 画像情報のリスト)
        """
        images_md = []
        images_info = []

        if not hasattr(sheet, '_images') or not sheet._images:
            return images_md, images_info

        output_dir = self.config.get('output_dir', 'images')

        # 出力ディレクトリの作成
        if not os.path.exists(output_dir):
            os.makedirs(output_dir)

        for img_idx, img in enumerate(sheet._images):
            try:
                self.image_counter += 1

                # 画像ファイル名の生成
                image_format = self.config.get('image_format', 'png')
                image_filename = f"chart_{self.image_counter:03d}.{image_format}"
                image_path = os.path.join(output_dir, image_filename)

                # 画像を保存
                self._save_image(img, image_path)

                # Markdown形式の画像参照を生成
                title = getattr(img, 'name', None) or f"Image {self.image_counter}"
                md_image = f"![{title}](./{image_path})"

                # 画像説明の生成（設定により）
                if self.config.get('generate_image_description', False):
                    description = self._generate_image_description(img)
                    if description:
                        md_image += f"\n\n{description}"

                images_md.append(md_image)
                images_info.append({
                    'index': self.image_counter,
                    'filename': image_filename,
                    'path': image_path,
                    'title': title
                })

            except Exception as e:
                print(f"画像抽出エラー: {str(e)}")
                continue

        return images_md, images_info

    def _save_image(self, img, output_path: str):
        """画像を保存"""
        try:
            # openpyxlの画像オブジェクトからPIL Imageに変換
            if hasattr(img, '_data'):
                # 画像データを取得
                image_data = img._data()
                pil_image = Image.open(io.BytesIO(image_data))

                # 最大サイズの制限（設定により）
                max_size = tuple(self.config.get('max_size', [1920, 1080]))
                if pil_image.size[0] > max_size[0] or pil_image.size[1] > max_size[1]:
                    pil_image.thumbnail(max_size, Image.LANCZOS)

                # ファイル形式に応じて保存
                image_format = self.config.get('image_format', 'png').upper()
                if image_format == 'JPG':
                    image_format = 'JPEG'

                pil_image.save(output_path, format=image_format)

        except Exception as e:
            print(f"画像保存エラー: {str(e)}")
            raise

    def _generate_image_description(self, img) -> str:
        """画像の説明を生成（基本的な情報のみ）"""
        description_parts = ["【画像情報】"]

        # 画像名
        if hasattr(img, 'name') and img.name:
            description_parts.append(f"- 名前: {img.name}")

        # 画像サイズ
        if hasattr(img, 'width') and hasattr(img, 'height'):
            description_parts.append(f"- サイズ: {img.width} x {img.height}")

        if len(description_parts) > 1:
            return '\n'.join(description_parts)

        return ""

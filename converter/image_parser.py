"""
Image Parser - 画像・グラフ・図形抽出ロジック
"""

import os
from typing import List, Tuple, Dict, Any
from PIL import Image
import io


class ImageParser:
    """画像・グラフ・図形抽出クラス"""

    def __init__(self, config: Dict[str, Any]):
        self.config = config
        self.image_counter = 0
        self.shape_counter = 0

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
                    'title': title,
                    'type': 'image'
                })

            except Exception as e:
                print(f"画像抽出エラー: {str(e)}")
                continue

        return images_md, images_info

    def extract_shapes(self, sheet) -> Tuple[List[str], List[Dict[str, Any]]]:
        """
        シートから図形とその中のテキストを抽出（テキストボックスを含む）

        Args:
            sheet: openpyxlのWorksheetオブジェクト

        Returns:
            (Markdown形式の図形情報リスト, 図形情報のリスト)
        """
        shapes_md = []
        shapes_info = []

        # openpyxlの図形オブジェクトにアクセス（_drawingを使用）
        if not hasattr(sheet, '_drawing') or not sheet._drawing:
            return shapes_md, shapes_info

        try:
            drawing = sheet._drawing

            # すべてのアンカータイプをチェック（テキストボックスはどのアンカータイプでも存在する可能性がある）
            anchor_lists = []

            if hasattr(drawing, 'twoCellAnchor') and drawing.twoCellAnchor:
                anchor_lists.append(('twoCellAnchor', drawing.twoCellAnchor))

            if hasattr(drawing, 'oneCellAnchor') and drawing.oneCellAnchor:
                anchor_lists.append(('oneCellAnchor', drawing.oneCellAnchor))

            if hasattr(drawing, 'absoluteAnchor') and drawing.absoluteAnchor:
                anchor_lists.append(('absoluteAnchor', drawing.absoluteAnchor))

            # すべてのアンカーから図形を抽出
            for anchor_type, anchors in anchor_lists:
                for anchor in anchors:
                    try:
                        # 図形（sp）を取得
                        shape = None
                        if hasattr(anchor, 'sp') and anchor.sp:
                            shape = anchor.sp

                        if not shape:
                            continue

                        self.shape_counter += 1

                        # 図形の基本情報
                        shape_data = {
                            'index': self.shape_counter,
                            'type': 'shape',
                            'anchor_type': anchor_type
                        }

                        # 図形名を取得
                        shape_name = f"Shape {self.shape_counter}"
                        if hasattr(shape, 'nvSpPr') and shape.nvSpPr:
                            if hasattr(shape.nvSpPr, 'cNvPr') and shape.nvSpPr.cNvPr:
                                name = getattr(shape.nvSpPr.cNvPr, 'name', None)
                                if name:
                                    shape_name = name

                        shape_data['name'] = shape_name

                        # 図形内のテキストを取得
                        shape_text = self._extract_text_from_shape(shape)

                        # テキストが取得できた場合のみMarkdownに追加
                        if shape_text:
                            shape_data['text'] = shape_text

                            # Markdown形式で出力
                            md_parts = [f"### 📐 {shape_name}"]
                            # テキストを引用として表示（複数行対応）
                            for line in shape_text.split('\n'):
                                if line.strip():
                                    md_parts.append(f"> {line}")

                            # 位置情報を追加
                            anchor_info = self._get_anchor_info(anchor)
                            if anchor_info:
                                shape_data['position'] = anchor_info
                                md_parts.append(f"\n**位置情報**: {anchor_info}")

                            md_shape = '\n'.join(md_parts)
                            shapes_md.append(md_shape)
                            shapes_info.append(shape_data)

                    except Exception as e:
                        print(f"図形抽出エラー（{anchor_type}）: {str(e)}")
                        import traceback
                        if self.config.get('verbose_logging', False):
                            traceback.print_exc()
                        continue

        except Exception as e:
            print(f"図形抽出全体エラー: {str(e)}")
            import traceback
            if self.config.get('verbose_logging', False):
                traceback.print_exc()

        return shapes_md, shapes_info

    def _extract_text_from_shape(self, shape) -> str:
        """
        図形からテキストを抽出

        Args:
            shape: 図形オブジェクト

        Returns:
            抽出されたテキスト
        """
        text_parts = []

        try:
            # txBodyからテキストを取得
            if hasattr(shape, 'txBody') and shape.txBody:
                txBody = shape.txBody

                # 段落（paragraph）のリストを取得
                paragraphs = []
                if hasattr(txBody, 'p'):
                    if isinstance(txBody.p, list):
                        paragraphs = txBody.p
                    else:
                        paragraphs = [txBody.p]

                # 各段落からテキストを抽出
                for paragraph in paragraphs:
                    if paragraph is None:
                        continue

                    paragraph_text = []

                    # run（テキストの塊）のリストを取得
                    runs = []
                    if hasattr(paragraph, 'r'):
                        if isinstance(paragraph.r, list):
                            runs = paragraph.r
                        else:
                            runs = [paragraph.r]

                    # 各runからテキストを抽出
                    for run in runs:
                        if run is None:
                            continue

                        if hasattr(run, 't') and run.t:
                            paragraph_text.append(str(run.t))

                    # 段落のテキストを結合
                    if paragraph_text:
                        text_parts.append(''.join(paragraph_text))

        except Exception as e:
            print(f"テキスト抽出エラー: {str(e)}")
            if self.config.get('verbose_logging', False):
                import traceback
                traceback.print_exc()

        # 段落を改行で結合
        return '\n'.join(text_parts) if text_parts else None

    def _get_anchor_info(self, anchor) -> str:
        """図形の位置情報を取得"""
        try:
            # twoCellAnchorの場合
            if hasattr(anchor, '_from'):
                from_cell = anchor._from
                if hasattr(from_cell, 'col') and hasattr(from_cell, 'row'):
                    from openpyxl.utils import get_column_letter
                    col_letter = get_column_letter(from_cell.col + 1)
                    return f"セル {col_letter}{from_cell.row + 1} 付近"
            # 別の方法でアンカー情報を取得
            elif hasattr(anchor, 'col') and hasattr(anchor, 'row'):
                from openpyxl.utils import get_column_letter
                col_letter = get_column_letter(anchor.col + 1)
                return f"セル {col_letter}{anchor.row + 1} 付近"
            return ""
        except Exception:
            return ""

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

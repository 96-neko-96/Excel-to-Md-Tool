"""
Image Parser - 画像・グラフ・図形抽出ロジック
"""

import os
from typing import List, Tuple, Dict, Any
from PIL import Image
import io
import zipfile
import xml.etree.ElementTree as ET


class ImageParser:
    """画像・グラフ・図形抽出クラス"""

    def __init__(self, config: Dict[str, Any]):
        self.config = config
        self.image_counter = 0
        self.shape_counter = 0
        self.workbook_path = None  # Excelファイルのパスを保存

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

    def extract_shapes(self, sheet, workbook_path=None) -> Tuple[List[str], List[Dict[str, Any]]]:
        """
        シートから図形とその中のテキストを抽出

        Args:
            sheet: openpyxlのWorksheetオブジェクト
            workbook_path: Excelファイルのパス（XML解析用）

        Returns:
            (Markdown形式の図形情報リスト, 図形情報のリスト)
        """
        shapes_md = []
        shapes_info = []

        # 1. コメント（ノート）を図形として抽出
        comments_md, comments_info = self._extract_comments(sheet)
        shapes_md.extend(comments_md)
        shapes_info.extend(comments_info)

        # 2. ExcelファイルのXML構造から図形を抽出（テキストボックスなど）
        if workbook_path and os.path.exists(workbook_path):
            xml_shapes_md, xml_shapes_info = self._extract_shapes_from_xml(workbook_path, sheet.title)
            shapes_md.extend(xml_shapes_md)
            shapes_info.extend(xml_shapes_info)

        return shapes_md, shapes_info

    def _extract_comments(self, sheet) -> Tuple[List[str], List[Dict[str, Any]]]:
        """
        シートからコメント（ノート）を抽出

        Args:
            sheet: openpyxlのWorksheetオブジェクト

        Returns:
            (Markdown形式のコメント情報リスト, コメント情報のリスト)
        """
        comments_md = []
        comments_info = []

        # シート内のすべてのセルをチェックしてコメントを探す
        for row in sheet.iter_rows():
            for cell in row:
                if cell.comment:
                    self.shape_counter += 1

                    # コメントデータ
                    comment_data = {
                        'index': self.shape_counter,
                        'type': 'comment',
                        'cell': cell.coordinate,
                        'text': cell.comment.text
                    }

                    # Markdown形式で出力
                    md_parts = [
                        f"### 💬 コメント ({cell.coordinate})",
                        f"> {cell.comment.text}"
                    ]

                    if cell.comment.author:
                        comment_data['author'] = cell.comment.author
                        md_parts.append(f"\n**作成者**: {cell.comment.author}")

                    comments_md.append('\n'.join(md_parts))
                    comments_info.append(comment_data)

        return comments_md, comments_info

    def _extract_shapes_from_xml(self, workbook_path: str, sheet_name: str) -> Tuple[List[str], List[Dict[str, Any]]]:
        """
        ExcelファイルのXML構造から図形（テキストボックスなど）を抽出

        Args:
            workbook_path: Excelファイルのパス
            sheet_name: シート名

        Returns:
            (Markdown形式の図形情報リスト, 図形情報のリスト)
        """
        shapes_md = []
        shapes_info = []

        try:
            # ExcelファイルはZIPファイルとして読み込める
            with zipfile.ZipFile(workbook_path, 'r') as zip_ref:
                # シートのIDを取得する必要があるため、workbook.xmlを読む
                workbook_xml = zip_ref.read('xl/workbook.xml')
                workbook_root = ET.fromstring(workbook_xml)

                # 名前空間の定義
                namespaces = {
                    'main': 'http://schemas.openxmlformats.org/spreadsheetml/2006/main',
                    'r': 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
                }

                # シート名からsheetIdを取得
                sheet_id = None
                for sheet_elem in workbook_root.findall('.//main:sheet', namespaces):
                    if sheet_elem.get('name') == sheet_name:
                        # r:idを取得
                        rid = sheet_elem.get('{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id')
                        if rid:
                            # rIdからシート番号を抽出（例: rId1 -> 1）
                            sheet_num = rid.replace('rId', '')
                            sheet_id = sheet_num
                            break

                if not sheet_id:
                    return shapes_md, shapes_info

                # 描画ファイルを探す（xl/drawings/drawing{n}.xml）
                # シートとdrawingの対応は xl/worksheets/_rels/sheet{n}.xml.rels で定義されている
                try:
                    rels_path = f'xl/worksheets/_rels/sheet{sheet_id}.xml.rels'
                    if rels_path in zip_ref.namelist():
                        rels_xml = zip_ref.read(rels_path)
                        rels_root = ET.fromstring(rels_xml)

                        # 描画ファイルのパスを取得
                        drawing_path = None
                        for rel in rels_root.findall('.//{http://schemas.openxmlformats.org/package/2006/relationships}Relationship'):
                            if 'drawing' in rel.get('Target', ''):
                                drawing_path = 'xl/' + rel.get('Target').replace('../', '')
                                break

                        if drawing_path and drawing_path in zip_ref.namelist():
                            # 描画XMLを読み込む
                            drawing_xml = zip_ref.read(drawing_path)
                            drawing_root = ET.fromstring(drawing_xml)

                            # 図形（テキストボックス）を抽出
                            shapes = self._parse_drawing_xml(drawing_root)
                            for shape in shapes:
                                self.shape_counter += 1
                                shape['index'] = self.shape_counter

                                # Markdown形式で出力
                                md_parts = [f"### 📐 {shape.get('name', 'Shape')}"]

                                if shape.get('text'):
                                    md_parts.append(f"> {shape['text']}")

                                shapes_md.append('\n'.join(md_parts))
                                shapes_info.append(shape)

                except KeyError:
                    # rels ファイルがない場合はスキップ
                    pass

        except Exception as e:
            print(f"XML図形抽出エラー: {str(e)}")

        return shapes_md, shapes_info

    def _parse_drawing_xml(self, drawing_root: ET.Element) -> List[Dict[str, Any]]:
        """
        描画XMLから図形情報を抽出

        Args:
            drawing_root: 描画XMLのルート要素

        Returns:
            図形情報のリスト
        """
        shapes = []

        # 名前空間の定義
        namespaces = {
            'xdr': 'http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing',
            'a': 'http://schemas.openxmlformats.org/drawingml/2006/main'
        }

        # テキストボックスや図形を探す
        for shape_elem in drawing_root.findall('.//xdr:sp', namespaces):
            shape_data = {'type': 'shape'}

            # 図形名を取得
            name_elem = shape_elem.find('.//xdr:nvSpPr/xdr:cNvPr', namespaces)
            if name_elem is not None:
                shape_data['name'] = name_elem.get('name', 'Shape')

            # テキストを取得
            text_parts = []
            for t_elem in shape_elem.findall('.//a:t', namespaces):
                if t_elem.text:
                    text_parts.append(t_elem.text)

            if text_parts:
                shape_data['text'] = '\n'.join(text_parts)

            if shape_data.get('name') or shape_data.get('text'):
                shapes.append(shape_data)

        return shapes

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

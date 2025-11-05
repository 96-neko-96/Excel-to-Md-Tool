"""
図形抽出機能のテストスクリプト
"""
import openpyxl
import os
import sys

# converterモジュールをインポート
from converter.core import ExcelToMarkdownConverter

def create_test_excel_with_shapes(output_path):
    """図形を含むテストExcelファイルを作成"""
    print("テスト用Excelファイルを作成中...")

    wb = openpyxl.Workbook()
    sheet = wb.active
    sheet.title = "テストシート"

    # 基本データを追加
    sheet['A1'] = "商品名"
    sheet['B1'] = "価格"
    sheet['A2'] = "商品A"
    sheet['B2'] = 1000
    sheet['A3'] = "商品B"
    sheet['B3'] = 2000

    # openpyxl.drawing.text を使ってテキストボックスを追加する方法
    # 注: openpyxlは図形の作成機能が限定的なので、手動で作成したファイルが必要

    wb.save(output_path)
    print(f"✓ テストファイル作成完了: {output_path}")
    print("注意: openpyxlでは図形の作成機能が限定的です。")
    print("      実際の図形を含むExcelファイルでテストしてください。")
    return output_path

def test_shape_extraction():
    """図形抽出機能のテスト"""
    print("\n=== 図形抽出機能テスト ===\n")

    # テストExcelファイル作成
    test_excel = "/tmp/test_shapes.xlsx"
    create_test_excel_with_shapes(test_excel)

    # 実際のExcelファイルがある場合はそちらを使用
    if len(sys.argv) > 1:
        test_excel = sys.argv[1]
        print(f"\n指定されたファイルを使用: {test_excel}")

    if not os.path.exists(test_excel):
        print(f"エラー: ファイルが見つかりません: {test_excel}")
        return

    # デバッグモードでExcelファイルを読み込み
    print(f"\nExcelファイルを読み込み中: {test_excel}")
    wb = openpyxl.load_workbook(test_excel, data_only=False)
    sheet = wb.active

    print(f"シート名: {sheet.title}")

    # 図形情報のデバッグ
    print("\n--- 図形情報のチェック ---")
    if hasattr(sheet, '_drawing') and sheet._drawing:
        print("✓ _drawing 属性が存在します")
        drawing = sheet._drawing

        # twoCellAnchor
        if hasattr(drawing, 'twoCellAnchor') and drawing.twoCellAnchor:
            print(f"  twoCellAnchor の数: {len(drawing.twoCellAnchor)}")
            for idx, anchor in enumerate(drawing.twoCellAnchor):
                print(f"    Anchor {idx + 1}:")
                if hasattr(anchor, 'sp') and anchor.sp:
                    print(f"      ✓ shape (sp) が存在")
                    sp = anchor.sp

                    # 名前
                    if hasattr(sp, 'nvSpPr') and sp.nvSpPr:
                        if hasattr(sp.nvSpPr, 'cNvPr') and sp.nvSpPr.cNvPr:
                            name = getattr(sp.nvSpPr.cNvPr, 'name', 'No name')
                            print(f"        名前: {name}")

                    # テキスト
                    if hasattr(sp, 'txBody') and sp.txBody:
                        print(f"        ✓ txBody が存在")
                        txBody = sp.txBody
                        if hasattr(txBody, 'p'):
                            paragraphs = txBody.p if isinstance(txBody.p, list) else [txBody.p]
                            for p_idx, paragraph in enumerate(paragraphs):
                                if paragraph and hasattr(paragraph, 'r'):
                                    runs = paragraph.r if isinstance(paragraph.r, list) else [paragraph.r]
                                    for run in runs:
                                        if run and hasattr(run, 't') and run.t:
                                            print(f"          テキスト: '{run.t}'")

        # oneCellAnchor
        if hasattr(drawing, 'oneCellAnchor') and drawing.oneCellAnchor:
            print(f"  oneCellAnchor の数: {len(drawing.oneCellAnchor)}")

        # absoluteAnchor
        if hasattr(drawing, 'absoluteAnchor') and drawing.absoluteAnchor:
            print(f"  absoluteAnchor の数: {len(drawing.absoluteAnchor)}")
    else:
        print("✗ _drawing 属性が存在しません（図形がありません）")

    # Converterを使って実際に変換
    print("\n--- Converterによる変換テスト ---")
    output_md = "/tmp/test_output.md"

    converter = ExcelToMarkdownConverter(
        extract_images=True,
        verbose_logging=True
    )

    try:
        result = converter.convert(test_excel, output_md)
        print(f"\n変換結果:")
        print(f"  シート数: {result['sheets_count']}")
        print(f"  テーブル数: {result['tables_count']}")
        print(f"  画像数: {result['images_count']}")

        # 図形情報を確認
        if hasattr(converter, 'sheets_data') and converter.sheets_data:
            total_shapes = sum(s.get('shapes_count', 0) for s in converter.sheets_data)
            print(f"  図形数: {total_shapes}")

            for sheet_data in converter.sheets_data:
                if sheet_data.get('shapes_count', 0) > 0:
                    print(f"\n  シート '{sheet_data['name']}' の図形情報:")
                    for shape in sheet_data.get('shapes', []):
                        print(f"    - {shape.get('name', 'Unknown')}: {shape.get('text', '(テキストなし)')[:50]}")

        # 生成されたMarkdownを表示
        print(f"\n生成されたMarkdownファイル: {output_md}")
        if os.path.exists(output_md):
            with open(output_md, 'r', encoding='utf-8') as f:
                content = f.read()

            # 図形セクションを探す
            if '📐' in content:
                print("\n✓ Markdownに図形情報が含まれています")
                lines = content.split('\n')
                for i, line in enumerate(lines):
                    if '📐' in line:
                        print(f"\n  図形セクション（行 {i+1}）:")
                        # 前後5行を表示
                        start = max(0, i-2)
                        end = min(len(lines), i+8)
                        for j in range(start, end):
                            print(f"    {lines[j]}")
                        if i > len(lines) - 10:
                            break
            else:
                print("\n✗ Markdownに図形情報が含まれていません")
                print("\n最初の100行を表示:")
                lines = content.split('\n')
                for i, line in enumerate(lines[:100]):
                    print(f"{i+1:3d}: {line}")

    except Exception as e:
        print(f"エラー: {e}")
        import traceback
        traceback.print_exc()

    print("\n=== テスト完了 ===")

if __name__ == "__main__":
    test_shape_extraction()

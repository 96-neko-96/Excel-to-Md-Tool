"""
ExcelファイルのZIP構造を調査
"""
import zipfile
import os
import sys
from xml.etree import ElementTree as ET

def analyze_excel_structure(excel_path):
    """ExcelファイルのZIP構造を解析"""
    if not os.path.exists(excel_path):
        print(f"ファイルが見つかりません: {excel_path}")
        return

    print(f"=== {excel_path} の構造解析 ===\n")

    try:
        with zipfile.ZipFile(excel_path, 'r') as zip_ref:
            # すべてのファイルをリスト
            print("【ファイル一覧】")
            for name in sorted(zip_ref.namelist()):
                print(f"  {name}")

            print("\n【drawing関連ファイルの内容】")
            # drawingファイルを探す
            drawing_files = [f for f in zip_ref.namelist() if 'drawing' in f.lower() and f.endswith('.xml')]

            if drawing_files:
                for drawing_file in drawing_files:
                    print(f"\n--- {drawing_file} ---")
                    try:
                        content = zip_ref.read(drawing_file).decode('utf-8')
                        # 最初の2000文字を表示
                        print(content[:2000])
                        if len(content) > 2000:
                            print("... (省略)")
                    except Exception as e:
                        print(f"読み込みエラー: {e}")
            else:
                print("  drawingファイルが見つかりません")

            print("\n【worksheet関連ファイル】")
            worksheet_files = [f for f in zip_ref.namelist() if 'worksheet' in f.lower() and f.endswith('.xml')]

            if worksheet_files:
                for ws_file in worksheet_files[:1]:  # 最初の1つだけ
                    print(f"\n--- {ws_file} (最初の3000文字) ---")
                    try:
                        content = zip_ref.read(ws_file).decode('utf-8')
                        print(content[:3000])
                        if len(content) > 3000:
                            print("... (省略)")
                    except Exception as e:
                        print(f"読み込みエラー: {e}")

    except zipfile.BadZipFile:
        print("エラー: 有効なZIPファイルではありません")
    except Exception as e:
        print(f"エラー: {e}")

if __name__ == "__main__":
    if len(sys.argv) > 1:
        analyze_excel_structure(sys.argv[1])
    else:
        # デフォルトのテストファイル
        test_file = "/tmp/test_shapes.xlsx"
        if os.path.exists(test_file):
            analyze_excel_structure(test_file)
        else:
            print("使用方法: python test_excel_structure.py <Excelファイルパス>")

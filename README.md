# Excel to Markdown Converter

RAGシステム用にExcelファイルをMarkdown形式に変換するWebアプリケーション

![Version](https://img.shields.io/badge/version-0.1.0-blue)
![Python](https://img.shields.io/badge/python-3.9+-green)
![License](https://img.shields.io/badge/license-MIT-yellow)

## 概要

複数シート、表、グラフを含む複雑なExcelファイルを、RAGシステムで利用しやすいMarkdown形式に変換するツールです。Streamlitベースの直感的なWebUIを提供し、シート間の参照関係を保持しながら、ブック単位で1つのMarkdownファイルに統合します。

## 主な機能

- ✅ **複数シート対応**: すべてのシートを1つのMarkdownファイルに統合
- ✅ **表の変換**: ExcelテーブルをMarkdown table形式に変換
- ✅ **画像抽出**: グラフや画像を抽出してファイル参照を生成
- ✅ **シート間参照**: 数式によるシート間の関連を保持
- ✅ **メタデータ生成**: RAGシステム用の詳細なメタデータを出力
- ✅ **目次自動生成**: シート構造から自動的に目次を作成
- ✅ **RAG最適化**: チャンクサイズ調整、キーワード抽出、階層構造の保持

## システム要件

- Python 3.9以上
- 推奨メモリ: 2GB以上

## インストール

### 1. リポジトリのクローン

```bash
git clone https://github.com/your-org/Excel-to-Md-Tool.git
cd Excel-to-Md-Tool
```

### 2. 仮想環境の作成（推奨）

```bash
python -m venv venv

# Windows
venv\Scripts\activate

# macOS/Linux
source venv/bin/activate
```

### 3. 依存パッケージのインストール

```bash
pip install -r requirements.txt
```

## 使い方

### Streamlit WebUIの起動

```bash
streamlit run app.py
```

ブラウザが自動的に開き、`http://localhost:8501` でアプリケーションにアクセスできます。

### 基本的な使用手順

1. **ファイルアップロード**: Excelファイル（.xlsx）をドラッグ&ドロップまたは選択
2. **設定調整**: サイドバーで変換オプションを設定
   - 目次生成のON/OFF
   - 画像抽出のON/OFF
   - チャンクサイズの調整（400-1500トークン）
3. **変換実行**: 「変換開始」ボタンをクリック
4. **プレビュー**: 変換結果を確認（Markdown/レンダリング/メタデータ）
5. **ダウンロード**: Markdownファイル、メタデータ、画像をダウンロード

## プロジェクト構造

```
Excel-to-Md-Tool/
├── app.py                      # Streamlit メインアプリ
├── converter/
│   ├── __init__.py
│   ├── core.py                 # ExcelToMarkdownConverter クラス
│   ├── sheet_parser.py         # シート解析ロジック
│   ├── table_parser.py         # 表検出・変換
│   ├── image_parser.py         # 画像・グラフ抽出
│   ├── markdown_generator.py  # Markdown生成
│   └── metadata_generator.py  # メタデータ生成
├── utils/
│   └── __init__.py
├── config.yaml                 # デフォルト設定
├── requirements.txt            # 依存パッケージ
└── README.md                   # このファイル
```

## 設定

### config.yaml

デフォルト設定は `config.yaml` で管理されています。主な設定項目：

```yaml
conversion:
  create_toc: true              # 目次生成
  extract_images: true          # 画像抽出
  generate_summary: false       # 表の要約生成

rag:
  chunk_size: 800               # トークン数
  chunk_overlap: 200            # オーバーラップ

table:
  format: "markdown"            # markdown / html
  max_columns: 20               # 最大列数

image:
  output_dir: "images"          # 出力ディレクトリ
  format: "png"                 # png / jpg
  max_size: [1920, 1080]        # 最大サイズ
```

## 出力形式

### Markdownファイル

```markdown
# ファイル名

**ファイル情報**
- シート数: N
- 変換日時: YYYY-MM-DD HH:MM:SS

## 目次
1. [シート1](#シート1)
2. [シート2](#シート2)

---

<a name="シート1"></a>
# 1. シート1

| 列1 | 列2 | 列3 |
|-----|-----|-----|
| 値1 | 値2 | 値3 |

> **関連シート:**
> - [シート2](#シート2) (参照先: セル B10)
```

### メタデータ (JSON)

```json
{
  "source_file": "example.xlsx",
  "converted_at": "2025-10-29T10:00:00",
  "total_sheets": 3,
  "sheets": [
    {
      "name": "Sheet1",
      "tables_count": 2,
      "images_count": 1,
      "keywords": ["売上", "予算", "実績"]
    }
  ],
  "cross_references": [
    {
      "from": "Sheet1!A5",
      "to": "Sheet2!B10",
      "type": "sum"
    }
  ],
  "rag_optimization": {
    "chunk_size": 800,
    "estimated_chunks": 15
  }
}
```

## 使用例

### 営業レポート

複数シートに分かれた売上データ、経費明細、予算表を1つのMarkdownファイルに統合し、RAGシステムで検索可能にする。

### プロジェクト管理資料

進捗表、リソース表、予算表を統合し、プロジェクト全体の情報を一元管理。

### データ分析レポート

データ集計表とグラフを含む分析結果を、検索可能な形式でドキュメント化。

## 技術スタック

- **Python 3.9+**
- **Streamlit** - WebUI
- **openpyxl** - Excel読み込み・書式解析
- **pandas** - データ処理・表変換
- **Pillow** - 画像処理
- **PyYAML** - 設定ファイル管理
- **tiktoken** - トークン数カウント（オプション）

## トラブルシューティング

### 画像が抽出されない

- Excelファイルに埋め込まれた画像のみ対応しています
- リンクされた外部画像は抽出されません

### 大きなファイルの処理が遅い

- チャンクサイズを大きくすることで処理速度が向上する場合があります
- メモリ不足の場合は、ファイルを分割することを検討してください

### エンコーディングエラー

- すべての出力はUTF-8エンコーディングです
- 特殊文字が含まれる場合は、Excelファイルを確認してください

## ロードマップ

### Phase 1: MVP（完了）
- ✅ 基本的な変換機能
- ✅ Streamlit UI
- ✅ 複数シート統合
- ✅ メタデータ生成

### Phase 2: 機能拡充（計画中）
- ⬜ LLMを使用した表の自動要約
- ⬜ グラフの説明自動生成
- ⬜ バッチ処理機能
- ⬜ 変換履歴管理

### Phase 3: エンタープライズ機能（将来）
- ⬜ API提供
- ⬜ ユーザー認証
- ⬜ クラウドストレージ連携

## ライセンス

MIT License

## 貢献

プルリクエストを歓迎します。大きな変更の場合は、まずissueを開いて変更内容を議論してください。

## 作成者

Claude Code - Excel to Markdown Converter

## 参考資料

- [Streamlit Documentation](https://docs.streamlit.io/)
- [openpyxl Documentation](https://openpyxl.readthedocs.io/)
- [pandas Documentation](https://pandas.pydata.org/docs/)

---

**Note**: このツールはRAGシステムとの統合を前提に設計されていますが、独立したツールとして使用できます。

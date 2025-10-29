"""
Excel to Markdown Converter - Streamlit Web UI
"""

import streamlit as st
from converter import ExcelToMarkdownConverter
import json
import os
import tempfile
import shutil


# ページ設定
st.set_page_config(
    page_title="Excel to Markdown Converter",
    page_icon="📊",
    layout="wide"
)

# セッション状態の初期化
if 'conversion_result' not in st.session_state:
    st.session_state.conversion_result = None

# タイトル
st.title("📊 Excel to Markdown Converter")
st.markdown("RAG用にExcelファイルをMarkdown形式に変換します")

# サイドバー: 設定パネル
with st.sidebar:
    st.header("⚙️ 変換設定")

    create_toc = st.checkbox("目次を生成", value=True, help="シート一覧の目次を自動生成します")
    extract_images = st.checkbox("画像を抽出", value=True, help="グラフや画像を抽出して保存します")
    generate_summary = st.checkbox("表の要約を生成", value=False, help="各テーブルの要約情報を追加します")
    chunk_size = st.slider(
        "チャンクサイズ (トークン)",
        min_value=400,
        max_value=1500,
        value=800,
        step=50,
        help="RAGシステム用のチャンクサイズ"
    )

    st.markdown("---")
    st.markdown("### 📖 使い方")
    st.markdown("""
    1. Excelファイルをアップロード
    2. 設定を確認・調整
    3. 変換ボタンをクリック
    4. プレビューを確認
    5. 結果をダウンロード
    """)

    st.markdown("---")
    st.markdown("### ℹ️ 対応形式")
    st.markdown("""
    - Excel 2007以降 (.xlsx)
    - 複数シート対応
    - 表・画像・グラフ対応
    - シート間参照の保持
    """)

# メインエリア
uploaded_file = st.file_uploader(
    "Excelファイルを選択してください",
    type=['xlsx'],
    help="複数シートを含むExcelファイルに対応しています"
)

if uploaded_file is not None:
    # ファイル情報表示
    col1, col2 = st.columns([3, 1])
    with col1:
        st.info(f"📁 **ファイル名:** {uploaded_file.name}")
    with col2:
        st.metric("サイズ", f"{uploaded_file.size / 1024:.1f} KB")

    # 変換ボタン
    if st.button("🔄 変換開始", type="primary", use_container_width=True):
        # 一時ディレクトリの作成
        temp_dir = tempfile.mkdtemp()

        try:
            with st.spinner("変換処理中..."):
                # 一時ファイル保存
                temp_input = os.path.join(temp_dir, uploaded_file.name)
                with open(temp_input, "wb") as f:
                    f.write(uploaded_file.getbuffer())

                # プログレスバー
                progress_bar = st.progress(0)
                status_text = st.empty()

                # 変換実行
                status_text.text("📖 Excelファイルを読み込み中...")
                progress_bar.progress(20)

                converter = ExcelToMarkdownConverter(
                    chunk_size=chunk_size,
                    create_toc=create_toc,
                    extract_images=extract_images,
                    generate_summary=generate_summary,
                    output_dir=os.path.join(temp_dir, 'images')
                )

                status_text.text("🔄 シートを変換中...")
                progress_bar.progress(40)

                # 出力ファイルパス
                output_filename = os.path.splitext(uploaded_file.name)[0] + '.md'
                output_file = os.path.join(temp_dir, output_filename)

                result = converter.convert(temp_input, output_file)

                status_text.text("📝 Markdownファイルを生成中...")
                progress_bar.progress(80)

                # 結果をセッションに保存
                with open(output_file, 'r', encoding='utf-8') as f:
                    md_content = f.read()

                # 画像ファイルの読み込み（あれば）
                images_data = {}
                images_dir = os.path.join(temp_dir, 'images')
                if os.path.exists(images_dir):
                    for img_file in os.listdir(images_dir):
                        img_path = os.path.join(images_dir, img_file)
                        with open(img_path, 'rb') as f:
                            images_data[img_file] = f.read()

                st.session_state.conversion_result = {
                    'md_content': md_content,
                    'metadata': result['metadata'],
                    'stats': result,
                    'original_filename': uploaded_file.name,
                    'images': images_data
                }

                progress_bar.progress(100)
                status_text.text("✅ 変換完了！")

        except Exception as e:
            st.error(f"❌ エラーが発生しました: {str(e)}")
            st.exception(e)

        finally:
            # 一時ディレクトリの削除
            try:
                shutil.rmtree(temp_dir)
            except:
                pass

# 変換結果の表示
if st.session_state.conversion_result:
    result = st.session_state.conversion_result

    st.markdown("---")
    st.success("✅ 変換が完了しました！")

    # 統計情報
    st.subheader("📊 変換結果サマリー")
    col1, col2, col3, col4 = st.columns(4)
    with col1:
        st.metric("シート数", result['stats']['sheets_count'])
    with col2:
        st.metric("テーブル数", result['stats']['tables_count'])
    with col3:
        st.metric("画像数", result['stats'].get('images_count', 0))
    with col4:
        st.metric("推奨チャンク数", result['stats']['estimated_chunks'])

    # プレビュー
    st.markdown("---")
    st.subheader("📄 変換結果プレビュー")

    # タブで表示切替
    tab1, tab2, tab3 = st.tabs(["📝 Markdown", "👁️ プレビュー", "📋 メタデータ"])

    with tab1:
        # Markdownのコード表示（最初の3000文字）
        preview_length = 3000
        md_preview = result['md_content'][:preview_length]
        if len(result['md_content']) > preview_length:
            md_preview += "\n\n... (以下省略)"

        st.code(md_preview, language="markdown", line_numbers=True)
        st.caption(f"全体の長さ: {len(result['md_content'])} 文字")

    with tab2:
        # Markdownのレンダリング表示
        preview_length = 3000
        md_preview = result['md_content'][:preview_length]
        if len(result['md_content']) > preview_length:
            md_preview += "\n\n*... (以下省略)*"

        st.markdown(md_preview)

    with tab3:
        # メタデータのJSON表示
        st.json(result['metadata'])

    # ダウンロードセクション
    st.markdown("---")
    st.subheader("💾 ダウンロード")

    col1, col2, col3 = st.columns(3)

    with col1:
        md_filename = os.path.splitext(result['original_filename'])[0] + '.md'
        st.download_button(
            label="📄 Markdownファイル",
            data=result['md_content'],
            file_name=md_filename,
            mime="text/markdown",
            use_container_width=True
        )

    with col2:
        metadata_json = json.dumps(result['metadata'], ensure_ascii=False, indent=2)
        metadata_filename = os.path.splitext(result['original_filename'])[0] + '_metadata.json'
        st.download_button(
            label="📋 メタデータ",
            data=metadata_json,
            file_name=metadata_filename,
            mime="application/json",
            use_container_width=True
        )

    with col3:
        # 画像ファイルがあれば、ZIPでダウンロード
        if result.get('images'):
            import zipfile
            from io import BytesIO

            zip_buffer = BytesIO()
            with zipfile.ZipFile(zip_buffer, 'w', zipfile.ZIP_DEFLATED) as zip_file:
                for img_name, img_data in result['images'].items():
                    zip_file.writestr(f"images/{img_name}", img_data)

            zip_buffer.seek(0)
            zip_filename = os.path.splitext(result['original_filename'])[0] + '_images.zip'

            st.download_button(
                label="🖼️ 画像ファイル",
                data=zip_buffer.getvalue(),
                file_name=zip_filename,
                mime="application/zip",
                use_container_width=True
            )

else:
    # 初期画面
    st.info("👆 Excelファイルをアップロードして変換を開始してください")

    # 機能説明
    with st.expander("📌 このツールについて"):
        st.markdown("""
        ### 🎯 主な機能

        - ✅ **複数シート対応**: すべてのシートを1つのMarkdownファイルに統合
        - ✅ **表の変換**: ExcelテーブルをMarkdown table形式に変換
        - ✅ **画像抽出**: グラフや画像を抽出してファイル参照を生成
        - ✅ **シート間参照**: 数式によるシート間の関連を保持
        - ✅ **メタデータ生成**: RAGシステム用の詳細なメタデータを出力
        - ✅ **目次自動生成**: シート構造から自動的に目次を作成

        ### 📋 使用例

        1. **営業レポート**: 複数シートの売上データや経費明細を統合
        2. **プロジェクト資料**: 進捗表、予算表、リソース表を一元化
        3. **分析資料**: データ集計とグラフを含む分析結果のドキュメント化

        ### 🔧 RAG最適化

        - チャンクサイズの調整可能
        - キーワード自動抽出
        - セクション階層の保持
        - ベクトルDB登録用メタデータ

        ### 📚 技術スタック

        - Python 3.9+
        - Streamlit (WebUI)
        - openpyxl (Excel処理)
        - pandas (データ変換)
        """)

# フッター
st.markdown("---")
st.markdown(
    """
    <div style='text-align: center; color: gray; font-size: 0.8em;'>
    Excel to Markdown Converter v0.1.0 | RAG Optimized
    </div>
    """,
    unsafe_allow_html=True
)

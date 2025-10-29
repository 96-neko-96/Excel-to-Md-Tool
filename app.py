"""
Excel to Markdown Converter - Streamlit Web UI (Phase 2)
"""

import streamlit as st
from converter import ExcelToMarkdownConverter
from utils.presets import PresetManager
from utils.batch_processor import BatchProcessor
from utils.history import HistoryManager
import json
import os
import tempfile
import shutil
from datetime import datetime


# ページ設定
st.set_page_config(
    page_title="Excel to Markdown Converter",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded"
)

# セッション状態の初期化
if 'conversion_result' not in st.session_state:
    st.session_state.conversion_result = None
if 'preset_manager' not in st.session_state:
    st.session_state.preset_manager = PresetManager()
if 'history_manager' not in st.session_state:
    st.session_state.history_manager = HistoryManager()
if 'current_preset' not in st.session_state:
    st.session_state.current_preset = "デフォルト"

# ヘッダー
st.title("📊 Excel to Markdown Converter v2.0")
st.markdown("**Phase 2機能搭載:** RAG最適化、バッチ処理、プリセット管理、変換履歴")

# タブで機能を分離
tab1, tab2, tab3, tab4 = st.tabs(["🔄 単一ファイル変換", "📦 バッチ処理", "📜 変換履歴", "⚙️ 設定管理"])

# =============================================================================
# タブ1: 単一ファイル変換
# =============================================================================
with tab1:
    # サイドバー: 設定パネル
    with st.sidebar:
        st.header("⚙️ 変換設定")

        # プリセット選択
        preset_names = st.session_state.preset_manager.get_preset_names()
        selected_preset = st.selectbox(
            "プリセット設定",
            preset_names,
            index=preset_names.index(st.session_state.current_preset) if st.session_state.current_preset in preset_names else 0,
            help="保存済みの設定プリセットから選択"
        )

        # プリセットが変更された場合
        if selected_preset != st.session_state.current_preset:
            st.session_state.current_preset = selected_preset

        # プリセットの設定を読み込み
        preset_config = st.session_state.preset_manager.get_preset(selected_preset)

        st.markdown("---")
        st.markdown("### 📝 詳細設定")

        create_toc = st.checkbox("目次を生成", value=preset_config.get('create_toc', True), help="シート一覧の目次を自動生成します")
        extract_images = st.checkbox("画像を抽出", value=preset_config.get('extract_images', True), help="グラフや画像を抽出して保存します")
        generate_summary = st.checkbox("表の要約を生成", value=preset_config.get('generate_summary', False), help="各テーブルの要約情報を追加します")
        show_formulas = st.checkbox("数式を備考として表示", value=preset_config.get('show_formulas', True), help="Excelの数式を備考欄に表示します")

        chunk_size = st.slider(
            "チャンクサイズ (トークン)",
            min_value=400,
            max_value=1500,
            value=preset_config.get('chunk_size', 800),
            step=50,
            help="RAGシステム用のチャンクサイズ"
        )

        st.markdown("---")

        # 現在の設定を新しいプリセットとして保存
        with st.expander("💾 現在の設定を保存"):
            new_preset_name = st.text_input("プリセット名", value="")
            new_preset_desc = st.text_area("説明", value="")
            if st.button("保存", use_container_width=True):
                if new_preset_name:
                    try:
                        st.session_state.preset_manager.add_preset(
                            new_preset_name,
                            {
                                'chunk_size': chunk_size,
                                'create_toc': create_toc,
                                'extract_images': extract_images,
                                'generate_summary': generate_summary,
                                'show_formulas': show_formulas
                            },
                            new_preset_desc
                        )
                        st.success(f"プリセット '{new_preset_name}' を保存しました")
                        st.rerun()
                    except Exception as e:
                        st.error(f"保存エラー: {e}")
                else:
                    st.warning("プリセット名を入力してください")

        st.markdown("---")
        st.markdown("### 📖 使い方")
        st.markdown("""
        1. Excelファイルをアップロード
        2. プリセットまたは詳細設定を調整
        3. 変換ボタンをクリック
        4. プレビューを確認
        5. 結果をダウンロード
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
                        show_formulas=show_formulas,
                        output_dir=os.path.join(temp_dir, 'images')
                    )

                    status_text.text("🔄 シートを変換中...")
                    progress_bar.progress(40)

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

                    # 履歴に追加
                    st.session_state.history_manager.add_record({
                        'input_file': uploaded_file.name,
                        'output_file': output_filename,
                        'sheets_count': result['sheets_count'],
                        'tables_count': result['tables_count'],
                        'images_count': result['images_count'],
                        'estimated_chunks': result['estimated_chunks'],
                        'preset_used': selected_preset
                    })

                    progress_bar.progress(100)
                    status_text.text("✅ 変換完了！")

            except Exception as e:
                st.error(f"❌ エラーが発生しました: {str(e)}")
                st.exception(e)

            finally:
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

        tab_md, tab_preview, tab_meta = st.tabs(["📝 Markdown", "👁️ プレビュー", "📋 メタデータ"])

        with tab_md:
            preview_length = 3000
            md_preview = result['md_content'][:preview_length]
            if len(result['md_content']) > preview_length:
                md_preview += "\n\n... (以下省略)"

            st.code(md_preview, language="markdown", line_numbers=True)
            st.caption(f"全体の長さ: {len(result['md_content'])} 文字")

        with tab_preview:
            preview_length = 3000
            md_preview = result['md_content'][:preview_length]
            if len(result['md_content']) > preview_length:
                md_preview += "\n\n*... (以下省略)*"

            st.markdown(md_preview)

        with tab_meta:
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
        st.info("👆 Excelファイルをアップロードして変換を開始してください")

# =============================================================================
# タブ2: バッチ処理
# =============================================================================
with tab2:
    st.header("📦 バッチ処理 - 複数ファイル一括変換")

    uploaded_files = st.file_uploader(
        "複数のExcelファイルを選択してください",
        type=['xlsx'],
        accept_multiple_files=True,
        help="複数のファイルを一度に変換できます"
    )

    if uploaded_files:
        st.info(f"📁 {len(uploaded_files)} ファイル選択済み")

        # バッチ変換設定
        with st.expander("⚙️ バッチ変換設定"):
            batch_preset = st.selectbox(
                "使用するプリセット",
                st.session_state.preset_manager.get_preset_names(),
                help="すべてのファイルに適用する設定"
            )

        # バッチ変換実行
        if st.button("🚀 バッチ変換開始", type="primary", use_container_width=True):
            temp_dir = tempfile.mkdtemp()
            output_dir = os.path.join(temp_dir, 'output')

            try:
                # プリセット設定を取得
                batch_config = st.session_state.preset_manager.get_preset(batch_preset)

                # バッチプロセッサーの初期化
                processor = BatchProcessor(**batch_config)

                # ファイルを一時ディレクトリに保存
                input_files = []
                for uploaded_file in uploaded_files:
                    temp_file = os.path.join(temp_dir, uploaded_file.name)
                    with open(temp_file, 'wb') as f:
                        f.write(uploaded_file.getbuffer())
                    input_files.append(temp_file)

                # プログレスバー
                progress_bar = st.progress(0)
                status_text = st.empty()

                def progress_callback(current, total, filename):
                    progress = int((current / total) * 100)
                    progress_bar.progress(progress)
                    status_text.text(f"変換中 ({current}/{total}): {filename}")

                # バッチ処理実行
                results = processor.process_files(input_files, output_dir, progress_callback)

                # 結果の表示
                summary = processor.get_summary()

                st.success(f"✅ バッチ変換完了: {summary['success']}/{summary['total']} ファイル成功")

                # サマリー表示
                col1, col2, col3, col4 = st.columns(4)
                with col1:
                    st.metric("成功", summary['success'])
                with col2:
                    st.metric("失敗", summary['failed'])
                with col3:
                    st.metric("総シート数", summary['total_sheets'])
                with col4:
                    st.metric("総テーブル数", summary['total_tables'])

                # 詳細結果
                with st.expander("📋 詳細結果"):
                    for r in results:
                        if r['status'] == 'success':
                            st.success(f"✅ {os.path.basename(r['input_file'])}")
                        else:
                            st.error(f"❌ {os.path.basename(r['input_file'])}: {r.get('error_message', '不明なエラー')}")

                # ダウンロード（ZIP化）
                import zipfile
                from io import BytesIO

                zip_buffer = BytesIO()
                with zipfile.ZipFile(zip_buffer, 'w', zipfile.ZIP_DEFLATED) as zip_file:
                    for root, dirs, files in os.walk(output_dir):
                        for file in files:
                            file_path = os.path.join(root, file)
                            arcname = os.path.relpath(file_path, output_dir)
                            zip_file.write(file_path, arcname)

                zip_buffer.seek(0)

                st.download_button(
                    label="📦 すべての変換結果をダウンロード (ZIP)",
                    data=zip_buffer.getvalue(),
                    file_name=f"batch_conversion_{datetime.now().strftime('%Y%m%d_%H%M%S')}.zip",
                    mime="application/zip",
                    use_container_width=True
                )

            except Exception as e:
                st.error(f"❌ バッチ処理エラー: {e}")
                st.exception(e)

            finally:
                try:
                    shutil.rmtree(temp_dir)
                except:
                    pass

# =============================================================================
# タブ3: 変換履歴
# =============================================================================
with tab3:
    st.header("📜 変換履歴")

    # 統計情報
    stats = st.session_state.history_manager.get_statistics()

    col1, col2, col3, col4 = st.columns(4)
    with col1:
        st.metric("総変換回数", stats['total_conversions'])
    with col2:
        st.metric("総シート数", stats['total_sheets'])
    with col3:
        st.metric("総テーブル数", stats['total_tables'])
    with col4:
        st.metric("平均シート数", stats['average_sheets'])

    st.markdown("---")

    # 履歴表示
    recent_history = st.session_state.history_manager.get_recent(20)

    if recent_history:
        for record in recent_history:
            with st.expander(f"📄 {record.get('input_file', '不明')} - {record.get('timestamp', '')[:19]}"):
                col1, col2 = st.columns(2)
                with col1:
                    st.write("**入力ファイル:**", record.get('input_file', '不明'))
                    st.write("**出力ファイル:**", record.get('output_file', '不明'))
                    st.write("**使用プリセット:**", record.get('preset_used', '不明'))
                with col2:
                    st.write("**シート数:**", record.get('sheets_count', 0))
                    st.write("**テーブル数:**", record.get('tables_count', 0))
                    st.write("**画像数:**", record.get('images_count', 0))
                    st.write("**推奨チャンク数:**", record.get('estimated_chunks', 0))

        # 履歴クリア
        if st.button("🗑️ 履歴をクリア", type="secondary"):
            st.session_state.history_manager.clear_all()
            st.success("履歴をクリアしました")
            st.rerun()
    else:
        st.info("変換履歴がありません")

# =============================================================================
# タブ4: 設定管理
# =============================================================================
with tab4:
    st.header("⚙️ 設定管理")

    # プリセット一覧
    st.subheader("📚 プリセット一覧")

    preset_names = st.session_state.preset_manager.get_preset_names()

    for preset_name in preset_names:
        preset = st.session_state.preset_manager.get_preset(preset_name)

        with st.expander(f"⚙️ {preset_name}"):
            st.write("**説明:**", preset.get('description', '説明なし'))
            st.write("**設定内容:**")
            st.json({k: v for k, v in preset.items() if k not in ['description', 'created_at', 'updated_at']})

            # デフォルトプリセット以外は削除可能
            if preset_name not in ["デフォルト", "RAG最適化", "完全変換", "軽量版"]:
                if st.button(f"🗑️ {preset_name} を削除", key=f"delete_{preset_name}"):
                    st.session_state.preset_manager.delete_preset(preset_name)
                    st.success(f"プリセット '{preset_name}' を削除しました")
                    st.rerun()

# フッター
st.markdown("---")
st.markdown(
    """
    <div style='text-align: center; color: gray; font-size: 0.8em;'>
    Excel to Markdown Converter v2.0 (Phase 2) | RAG Optimized | Batch Processing | History Management
    </div>
    """,
    unsafe_allow_html=True
)

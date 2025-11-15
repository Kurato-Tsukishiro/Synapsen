import shutil
import datetime
import io
import re
import sqlite3
from pathlib import Path
from pypdf import PdfReader, PdfWriter
import fitz  # PyMuPDF (プレースホルダー生成用に追加)

# グラフ出力のためにGraphManagerを利用
from graph_manager import GraphManager

import logging
logger = logging.getLogger(__name__)


class ExportManager:
    def __init__(self, config_data):
        """
        Args:
            config_data (dict): アプリケーション設定 (key_icons, colors等を含む)
        """
        self.config = config_data

    def execute_export(
        self,
        target_df,
        query_text,
        export_folder_path,
        mode="search_results",
        include_pdf=False,
        loaded_db_path=None,
        pdf_root_folder=None
    ):
        """
        検索結果のエクスポート処理を実行する。
        """
        export_path = Path(export_folder_path)
        timestamp = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
        folder_suffix = "Selected" if mode == "selected_items" else "Search"
        file_name = f"Synapsen_Export_{folder_suffix}_{timestamp}"
        final_export_dir = export_path / file_name

        try:
            final_export_dir.mkdir(parents=True, exist_ok=True)

            # 1. メタデータ (TXT)
            self._save_meta_txt(
                final_export_dir, mode, len(target_df), query_text)

            # 2. CSV
            self._save_csv(final_export_dir, target_df)

            # 3. 本文テキスト (個別TXT)
            self._save_full_text(
                final_export_dir, target_df, loaded_db_path
            )

            # 4. グラフ (HTML)
            graph_path = final_export_dir / "relation_graph.html"
            GraphManager.generate_graph_html(
                target_df,
                self.config.get('key_icons', {}),
                self.config.get('key_colors', {}),
                loaded_db_path,
                pdf_root_folder,
                output_path=graph_path
            )

            # 5. 統合PDF (オプション)
            if include_pdf:
                pdf_save_path = final_export_dir / "Merged_Notes.pdf"
                self.merge_pdf(
                    target_df,
                    pdf_save_path, pdf_root_folder,
                    loaded_db_path
                )

            return True, final_export_dir

        except Exception as e:
            # エラー時は途中作成のフォルダを削除
            if final_export_dir.exists():
                try:
                    shutil.rmtree(final_export_dir)
                except Exception as e_del:
                    logger.error(
                        f"エラー後のエクスポートフォルダ削除に失敗: {e_del}",
                        extra={'sensitive': True}
                    )
                    pass
            raise e

    def merge_pdf(
            self, target_df,
            save_path, pdf_root_folder,
            loaded_db_path=None
    ):
        """
        DataFrame内のPDFを結合し、しおりを付けて保存する。
        統合PDF -> 元PDF -> フォールバック(代替ページ) の順で取得を試みる。
        """
        writer = PdfWriter()
        processed_count = 0
        current_page_index = 0

        for _, row in target_df.iterrows():
            # ノート情報の取得
            note_title = row.get('title', 'No Title')
            note_key = row.get('key', 'Unknown Key')
            date_str = row.get('date', '??????')

            if len(date_str) == 8:
                fmt_date = f"{date_str[:4]}/{date_str[4:6]}/{date_str[6:]}"
            else:
                fmt_date = date_str

            bookmark_title = f"{fmt_date} {note_title}"

            # --- 戦略 A: 統合PDF (月刊ノート) から取得 ---
            merged_filename = row.get('merged_pdf_filename')
            merged_start_page_str = row.get('merged_start_page')
            try:
                page_count = int(row.get('pages', 0))
            except Exception as e:
                logger.error(f"ページ数の取得エラー: {e}")
                page_count = 0

            pdf_source_path = None
            source_type = None  # "merged" or "original" or "fallback"
            start_p = 0

            # 1. 統合PDFのパスを確認
            if loaded_db_path and merged_filename:
                potential_path = Path(loaded_db_path).parent / merged_filename
                if (
                        potential_path.is_file()
                        and merged_start_page_str
                        and str(merged_start_page_str).isdigit()
                ):
                    pdf_source_path = potential_path
                    # DBは1始まり、pypdfは0始まりなので -1
                    start_p = int(merged_start_page_str) - 1
                    source_type = "merged"

            # 2. 統合PDFがなければ、元のPDFを確認
            if not pdf_source_path:
                filepath_str = row.get('filepath', '')
                if filepath_str:
                    potential_path = Path(filepath_str)
                    if not potential_path.is_absolute() and pdf_root_folder:
                        potential_path = Path(pdf_root_folder) / potential_path

                    if potential_path.is_file():
                        pdf_source_path = potential_path
                        start_p = 0
                        source_type = "original"

            # --- PDF結合処理 ---
            reader = None

            # ソースが見つかった場合
            if pdf_source_path:
                try:
                    reader = PdfReader(pdf_source_path)
                except Exception as e:
                    logger.error(
                        "[ExportManager] PDF read error "
                        f"({pdf_source_path.name}): {e}",
                        extra={'sensitive': True}
                    )
                    reader = None

            # ソースがない、または読み込みエラーの場合はフォールバック生成
            if reader is None:
                logger.info(
                    "[ExportManager] Creating fallback page for: "
                    f"{note_title}",
                    extra={'sensitive': True}
                )
                reader = self._create_fallback_page_reader(
                    note_title, fmt_date, note_key)
                source_type = "fallback"
                # フォールバックページは常に1ページとみなす
                start_p = 0
                page_count = 1

            # ページ追加実行
            if reader:
                try:
                    # しおり追加 (現在のページ位置)
                    writer.add_outline_item(bookmark_title, current_page_index)

                    if source_type == "merged":
                        # 統合PDF: 指定範囲だけ追加
                        pages_to_add = page_count if page_count > 0 else 1
                        for i in range(pages_to_add):
                            p_idx = start_p + i
                            if 0 <= p_idx < len(reader.pages):
                                writer.add_page(reader.pages[p_idx])
                                current_page_index += 1

                    elif source_type == "original":
                        # 元PDF: 全ページ追加
                        for page in reader.pages:
                            writer.add_page(page)
                            current_page_index += 1

                    elif source_type == "fallback":
                        # フォールバック: 生成された1ページを追加
                        writer.add_page(reader.pages[0])
                        current_page_index += 1

                    processed_count += 1

                except Exception as e:
                    logger.error(f"[ExportManager] PDF add page error: {e}")

        if processed_count > 0:
            with open(save_path, "wb") as f:
                writer.write(f)
            return True
        return False

    def _create_fallback_page_reader(self, title, date, key):
        """
        PyMuPDFを使用して「ファイルが見つかりません」というPDFページをオンメモリで生成し、
        pypdf.PdfReader オブジェクトとして返す。
        """
        try:
            doc = fitz.open()
            page = doc.new_page(width=595, height=842)   # A4 size

            # デザイン設定
            rect_header = fitz.Rect(50, 100, 545, 200)
            rect_body = fitz.Rect(50, 220, 545, 500)

            # ヘッダー（警告）
            page.insert_textbox(
                rect_header,
                "FILE NOT FOUND",
                fontsize=24,
                fontname="helv",
                color=(1, 0, 0),  # Red
                align=1  # Center
            )

            # 詳細情報
            info_text = (
                f"Original file is missing or inaccessible.\n\n"
                f"Date: {date}\n"
                f"Title: {title}\n"
                f"Key: {key}\n\n"
                f"This placeholder preserves the page order."
            )

            page.insert_textbox(
                rect_body,
                info_text,
                fontsize=14,
                fontname="helv",
                align=0  # Left
            )

            # メモリ上にPDFを保存
            pdf_bytes = doc.tobytes()
            doc.close()

            # BytesIOを経由してpypdfで読み込む
            return PdfReader(io.BytesIO(pdf_bytes))

        except Exception as e:
            logger.error(f"[ExportManager] Fallback generation failed: {e}")
            return None

    def _save_meta_txt(self, dir_path, mode, count, query):
        mode_str = '選択アイテムのみ' if mode == 'selected_items' else '検索結果全体'
        with open(dir_path / "export_meta.txt", 'w', encoding='utf-8') as f:
            f.write("Synapsen Nexus エクスポート\n")
            f.write(f"モード: {mode_str}\n")
            f.write(f"件数: {count} 件\n")
            f.write("=" * 30 + "\n")
            f.write(f"検索クエリ:\n{query}\n")

    def _save_csv(self, dir_path, df):
        csv_path = dir_path / "metadata.csv"
        df_export = df.drop(columns=['full_text'], errors='ignore')
        if 'tags' in df_export.columns:
            df_export['tags'] = df_export['tags'].apply(
                lambda x: ";".join(x) if isinstance(x, list) else str(x)
            )
        df_export.to_csv(csv_path, index=False, encoding='utf-8-sig')

    def _save_full_text(self, dir_path, df, loaded_db_path):
        """
        対象DataFrameのキーに基づき、DBからfull_textを取得して保存する。
        """
        text_dir = dir_path / "FullText_Contents"
        text_dir.mkdir(exist_ok=True)

        # 1. DBから full_text を取得するための準備
        keys_to_fetch = list(df['key'].dropna().unique())
        full_text_map = {}

        if keys_to_fetch and loaded_db_path:
            conn = None
            try:
                # 読み取り専用で接続
                conn = sqlite3.connect(
                    f"file:{loaded_db_path}?mode=ro", uri=True
                )
                cursor = conn.cursor()

                # プレースホルダを作成
                placeholders = ','.join('?' for _ in keys_to_fetch)
                sql = (
                    f"SELECT key, full_text FROM notes "
                    f"WHERE key IN ({placeholders})"
                )

                cursor.execute(sql, keys_to_fetch)

                # key -> full_text の辞書を作成
                for key, text in cursor.fetchall():
                    full_text_map[key] = text if text else ""

            except Exception as e:
                logger.error(f"[ExportManager] full_text の一括取得に失敗: {e}")
            finally:
                if conn:
                    conn.close()

        # 2. ファイルへの書き出し
        for _, row in df.iterrows():
            key = row.get('key')
            if key:
                text = full_text_map.get(key, '')

                # --- ここから整形処理 ---
                if text:
                    # 1. 文末記号の直後に改行を挿入
                    #    句点、疑問符などはそのまま改行 ([。？！」])
                    #    半角の疑問符と感嘆符及びピリオドが連続している場合、最後の記号の直後のみ改行 (\.+(?!\.))
                    text = re.sub(r'([。？！」]|\.+(?!\.))\n*', r'\1\n', text)

                    # 2. 連続する空白やタブを単一スペースに置換
                    text = re.sub(r'[ \t]+', ' ', text)

                    # 3. 連続する改行を2つまでに制限
                    text = re.sub(r'\n{3,}', '\n\n', text)
                # --- 整形処理ここまで ---

                with open(
                    text_dir / f"{key}.txt", 'w', encoding='utf-8'
                ) as f:
                    f.write(text)

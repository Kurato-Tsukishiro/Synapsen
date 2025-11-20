import shutil
import datetime
import io
import re
import sqlite3
from pathlib import Path
import pandas as pd
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

    def generate_moc_markdown(
            self,
            target_df: pd.DataFrame,
            save_path: str,
            loaded_db_path: str = None
    ) -> bool:
        """
        選択されたノート情報に基づき、MOC (Map of Content) 用のMarkdownファイルを生成する。

        データベースから最新の `memo` と `summary` を取得し、統合PDFまたは個別ファイルへのリンクを含む
        インデックスページを作成する。また、メモ内の内部リンク（`[[Link]]`）を抽出し、
        引用一覧セクションを生成する。

        Args:
            target_df (pd.DataFrame): 処理対象のノート情報を含むDataFrame。
                以下のカラムが含まれていることを想定:
                - 'key': ノートの一意なキー
                - 'title': ノートのタイトル
                - 'date', 'time': ソート用
                - 'filepath': 元ファイルのパス
                - 'merged_pdf_filename', 'merged_start_page': 統合PDF時のリンク用情報
                - 'commonplace_key': アイコン/色決定用のキー
            save_path (str | Path): 生成されたMarkdownファイルの保存先パス。
            loaded_db_path (str, optional): データ取得元のSQLiteデータベースのパス。
                指定された場合、`target_df` の値ではなくDBから最新の `memo`, `summary` を再取得して使用する。
                Defaults to None.

        Returns:
            bool: 生成に成功した場合は True、エラーが発生した場合は False を返す。
        """
        try:
            # --- 1. DBから必要なデータを一括取得するための準備 ---
            keys_in_moc = target_df['key'].dropna().tolist()
            db_data_map = {}      # key -> {'summary': ...}
            links_by_source = {}  # source_key -> [target_key, ...]
            all_linked_keys = set()
            title_map = {}        # key -> title

            conn = None
            if loaded_db_path and keys_in_moc:
                try:
                    # 読み取り専用でDB接続
                    conn = sqlite3.connect(
                        f"file:{loaded_db_path}?mode=ro",
                        uri=True
                    )
                    cursor = conn.cursor()

                    placeholders = ','.join('?' for _ in keys_in_moc)

                    # A. Summary の一括取得
                    sql_summary = (
                        "SELECT key, summary FROM notes "
                        f"WHERE key IN ({placeholders})"
                    )
                    cursor.execute(sql_summary, keys_in_moc)
                    for row in cursor.fetchall():
                        k, s = row
                        db_data_map[k] = {'summary': s if s else ""}

                    # B. 発リンク (Outgoing Links) の一括取得
                    # MOCを構成するノートからの発リンクのみを取得
                    sql_links = (
                        "SELECT source_key, target_key FROM note_links "
                        f"WHERE source_key IN ({placeholders})"
                    )
                    cursor.execute(sql_links, keys_in_moc)

                    for source_key, target_key in cursor.fetchall():
                        if source_key not in links_by_source:
                            links_by_source[source_key] = []
                        links_by_source[source_key].append(target_key)
                        all_linked_keys.add(target_key)  # リンク先もタイトルの取得対象に含める

                    # C. MOC内のキーとリンク先キーを合わせた全てのキーの集合
                    all_relevant_keys = set(keys_in_moc) | all_linked_keys

                    # D. 全ての関連ノートのタイトルを一括取得
                    title_placeholders = (
                        ','.join('?' for _ in all_relevant_keys)
                    )
                    sql_titles = (
                        "SELECT key, title FROM notes "
                        f"WHERE key IN ({title_placeholders})"
                    )
                    cursor.execute(sql_titles, list(all_relevant_keys))

                    for key, title in cursor.fetchall():
                        title_map[key] = title if title else "(タイトルなし)"

                except Exception as e:
                    logger.error(f"MOC生成時のDB取得エラー: {e}")
                finally:
                    if conn:
                        conn.close()
            # -------------------------------------------------------

            # タイトル生成
            timestamp = datetime.datetime.now().strftime("%Y-%m-%d")
            md_content = [
                f"# MOC (Map of Content)<br> ( {timestamp} )",
                "",
                "## 概要",
                "ここにこのMOCの概要や目的を記述してください。",
                "",
                "## 関連ノート一覧",
                ""
            ]

            # 日付・時刻順でソート
            sorted_df = target_df.sort_values(by=["date", "time"])

            # 設定からアイコンと色を取得
            key_icons = self.config.get('key_icons', {})
            key_colors = self.config.get('key_colors', {})

            # --- 第1部: 関連ノート一覧 ---
            for _, row in sorted_df.iterrows():
                key = row.get('key', '')
                title = row.get('title', 'No Title')

                # Summary
                summary_text = (
                    db_data_map.get(key, {}).get('summary', '').strip()
                )
                if not summary_text:
                    summary_text = "(概要なし)"

                # Index Key情報の取得と装飾
                cp_key_raw = str(row.get('commonplace_key', ''))
                cp_key_lower = cp_key_raw.lower()
                icon = key_icons.get(cp_key_lower, '')
                color = key_colors.get(cp_key_lower, '')
                display_content = icon if icon else cp_key_raw

                if color:
                    prefix_html = (
                        f'<span style="color: {color}">'
                        f'{display_content}'
                        '</span>'
                    )
                else:
                    prefix_html = display_content

                # リンク先の決定 (FilePath)
                filepath = row.get('filepath', '')
                original_filename = Path(filepath).name if filepath else ""

                merged_filename = row.get('merged_pdf_filename', '')
                merged_page = row.get('merged_start_page', '')

                link_target = "(File not found)"
                if merged_filename and merged_page:
                    # PDFのページ指定リンク形式: ./File.pdf#page=5
                    link_target = f"./{merged_filename}#page={merged_page}"

                # 2. なければ元ファイルを使用 (フォールバック)
                elif original_filename:
                    link_target = f"./{original_filename}"

                # ヘッダーにプレフィックスを統合
                item_header = f"- {prefix_html} [[{key}: {title}]]"

                # リスト項目の生成
                line = (
                    f"{item_header}\n"
                    f"  - 📂 {link_target}\n"
                    f"  - 💡 {summary_text}<br>"
                )
                md_content.append(line)

            # --- 改ページ ---
            md_content.append("")
            md_content.append('<div style="page-break-after: always;"></div>')
            md_content.append("")
            md_content.append("")

            # --- 第2部: 引用一覧 ---
            md_content.append("## 引用一覧")
            md_content.append("")

            has_references = False

            for _, row in sorted_df.iterrows():
                source_key = row.get('key', '')

                # リンクテーブルからターゲットキーを取得
                target_keys = links_by_source.get(source_key)

                if target_keys:
                    has_references = True
                    source_title = row.get('title', 'No Title')

                    # 見出し: 引用元のノート
                    md_content.append(f"### [[{source_key}: {source_title}]]")

                    # 引用リスト
                    for target_key in target_keys:
                        # title_map を使ってタイトルを取得し、リンク形式に整形
                        target_title = title_map.get(
                            target_key, f"(タイトル不明 / Key: {target_key})"
                        )
                        link_text = f"[[{target_key}: {target_title}]]"
                        md_content.append(f"- {link_text}")

                    md_content.append("")

            if not has_references:
                md_content.append("(引用リンクはありません)")

            md_content.append("")
            md_content.append("---")
            md_content.append(f"Generated by Synapsen Nexus on {timestamp}")

            # ファイル書き出し
            with open(save_path, 'w', encoding='utf-8') as f:
                f.write("\n".join(md_content))

            return True

        except Exception as e:
            logger.error(f"MOC生成エラー: {e}")
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

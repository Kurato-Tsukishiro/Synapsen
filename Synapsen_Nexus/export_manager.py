import shutil
import datetime
from pathlib import Path
from pypdf import PdfReader, PdfWriter

# グラフ出力のためにGraphManagerを利用
from graph_manager import GraphManager


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
            self._save_full_text(final_export_dir, target_df)

            # 4. グラフ (HTML)
            # 保存先パスを指定してGraphManagerを呼び出す
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
                self.merge_pdf(target_df, pdf_save_path, pdf_root_folder)

            return True, final_export_dir

        except Exception as e:
            # エラー時は途中作成のフォルダを削除
            if final_export_dir.exists():
                try:
                    shutil.rmtree(final_export_dir)
                except Exception as e_del:
                    print(f"エラー後のエクスポートフォルダ削除に失敗: {e_del}")
                    pass
            raise e

    def merge_pdf(self, target_df, save_path, pdf_root_folder):
        """
        DataFrame内のPDFを結合し、しおりを付けて保存する。
        """
        writer = PdfWriter()
        processed_count = 0
        current_page_index = 0

        for _, row in target_df.iterrows():
            filepath_str = row.get('filepath', '')
            if not filepath_str:
                continue

            pdf_path = Path(filepath_str)
            if not pdf_path.is_absolute() and pdf_root_folder:
                pdf_path = Path(pdf_root_folder) / pdf_path

            if pdf_path.is_file():
                try:
                    reader = PdfReader(pdf_path)

                    # しおり追加 (YYYY/MM/DD Title)
                    date_str = row.get('date', '??????')
                    if len(date_str) == 8:
                        yyyy = date_str[:4]
                        mm = date_str[4:6]
                        dd = date_str[6:]
                        fmt_date = f"{yyyy}/{mm}/{dd}"
                    else:
                        fmt_date = date_str

                    title = row.get('title', 'No Title')
                    bookmark_title = f"{fmt_date} {title}"
                    writer.add_outline_item(bookmark_title, current_page_index)

                    for page in reader.pages:
                        writer.add_page(page)
                        current_page_index += 1

                    processed_count += 1
                except Exception as e:
                    print(
                        "[ExportManager] PDF merge warning "
                        f"({pdf_path.name}): {e}"
                    )

        if processed_count > 0:
            with open(save_path, "wb") as f:
                writer.write(f)
            return True
        return False

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
        # リスト型のタグを文字列に戻す
        if 'tags' in df_export.columns:
            df_export['tags'] = df_export['tags'].apply(
                lambda x: ";".join(x) if isinstance(x, list) else str(x)
            )
        df_export.to_csv(csv_path, index=False, encoding='utf-8-sig')

    def _save_full_text(self, dir_path, df):
        text_dir = dir_path / "FullText_Contents"
        text_dir.mkdir(exist_ok=True)
        for _, row in df.iterrows():
            key = row.get('key')
            if key:
                text = row.get('full_text', '')
                # ファイル名はKeyを使用
                with open(text_dir / f"{key}.txt", 'w', encoding='utf-8') as f:
                    f.write(text)

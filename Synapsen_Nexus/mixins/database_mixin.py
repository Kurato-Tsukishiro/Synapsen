import sqlite3
import threading
import queue
import logging

import pandas as pd
from Synapsen_Nexus.utils import count_notes_from_db, fetch_notes_from_db
from Synapsen_Nexus.search_parser import parse_query_to_sql

logger = logging.getLogger("Nexus.DBMixin")


class NexusDatabaseMixin:
    """データベース操作と検索処理を担当するMixin"""

    def setup_database_variables(self):
        """DB関連変数の初期化"""
        self.loaded_db_path = None
        self.db_conn = None
        self._current_search_id = 0
        self._search_lock = threading.Lock()
        self._search_result_queue = queue.Queue()

        # 検索結果監視ループを開始
        self.after(100, self._poll_search_result)

    def load_db_from_path(self, filepath, key_to_redisplay=None):
        """DB接続を確立し、初期検索を行う"""
        try:
            if self.db_conn:
                self.db_conn.close()

            self.loaded_db_path = filepath
            # 読み取り専用接続
            self.db_conn = sqlite3.connect(f"file:{filepath}?mode=ro", uri=True)

            # 書き込み用接続でテーブルチェック (存在しなければ作成)
            self._ensure_tables_exist(filepath)

            # UIリセット & 初期検索
            self.current_page = 0
            self.refresh_unique_tags()
            self.perform_search()

            if key_to_redisplay:
                notes = fetch_notes_from_db(self.db_conn, "key = ?", [key_to_redisplay])
                if notes:
                    # Seriesではなく辞書として渡す
                    self.show_details(dict(notes[0]))
            else:
                self.clear_details()

        except Exception as e:
            logger.error(f"DB Load Error: {e}")
            from tkinter import messagebox

            messagebox.showerror("DBエラー", str(e))

    def _ensure_tables_exist(self, filepath):
        """必要なテーブルが存在するか確認・作成する"""
        conn_write = None
        try:
            conn_write = sqlite3.connect(filepath)
            cursor = conn_write.cursor()
            cursor.execute(
                """
                CREATE TABLE IF NOT EXISTS note_links (
                    source_key TEXT NOT NULL,
                    target_key TEXT NOT NULL,
                    PRIMARY KEY (source_key, target_key)
                )
            """
            )
            cursor.execute(
                "CREATE INDEX IF NOT EXISTS idx_target_key ON note_links (target_key)"
            )
            conn_write.commit()
        except Exception as e:
            logger.error(f"Table creation error: {e}")
        finally:
            if conn_write:
                conn_write.close()

    def perform_search(self, reset_page=True):
        """非同期検索を実行する"""
        if not self.db_conn:
            return

        if reset_page:
            self.current_page = 0

        # UIコンポーネントから値を取得
        user_query = self.search_entry.get().strip()
        include_full_text = self.fts_checkbox.get()

        # 除外タグ条件の構築
        exclusion_query = ""
        if self.exclude_tags_checkbox.get() == 1 and getattr(
            self, "exclude_tags_by_default", []
        ):
            parts = [f"-tag:{t}" for t in self.exclude_tags_by_default]
            exclusion_query = " ".join(parts)

        full_query_str = (
            f"({user_query}) AND ({exclusion_query})"
            if user_query and exclusion_query
            else (user_query or exclusion_query)
        )

        # フィルター条件
        selected_filters = [
            k for k, v in self.filter_checkboxes.items() if v.get() == "1"
        ]

        # SQL構築
        where_parts = []
        params = []

        if full_query_str:
            q_sql, q_params = parse_query_to_sql(full_query_str, include_full_text)
            if q_sql:
                where_parts.append(q_sql)
                params.extend(q_params)

        if selected_filters:
            placeholders = ",".join(["?"] * len(selected_filters))
            where_parts.append(f"commonplace_key IN ({placeholders})")
            params.extend(selected_filters)

        final_where = " AND ".join(where_parts) if where_parts else ""

        # 状態保存 (ページネーション用)
        self.current_where_clause = final_where
        self.current_params = params

        # スレッド起動ID発行
        with self._search_lock:
            self._current_search_id += 1
            search_id = self._current_search_id

        self.results_list.configure(label_text="検索中...")
        self.configure(cursor="watch")

        thread = threading.Thread(
            target=self._execute_search_worker,
            args=(
                search_id,
                self.loaded_db_path,
                final_where,
                params,
                self.current_page,
                self.items_per_page,
                self.sort_ascending,
            ),
            daemon=True,
        )
        thread.start()

    def _execute_search_worker(
        self, search_id, db_path, where, params, page, limit, asc
    ):
        """別スレッドで実行される検索処理"""
        try:
            conn = sqlite3.connect(f"file:{db_path}?mode=ro", uri=True)
            total = count_notes_from_db(conn, where, params)
            offset = page * limit
            rows = fetch_notes_from_db(conn, where, params, limit, offset, asc)
            conn.close()
            self._search_result_queue.put((search_id, (rows, total)))
        except Exception as e:
            logger.error(f"Search thread error: {e}")
            self._search_result_queue.put((search_id, ([], 0)))

    def _poll_search_result(self):
        """検索結果キューを監視する"""
        try:
            while True:
                search_id, result_data = self._search_result_queue.get_nowait()
                self._on_search_complete(search_id, result_data)
        except queue.Empty:
            pass
        finally:
            self.after(100, self._poll_search_result)

    def _on_search_complete(self, search_id, result_data):
        """検索完了時のUI更新"""
        rows, total_count = result_data
        self.configure(cursor="")

        with self._search_lock:
            if search_id != self._current_search_id:
                return

        self.filtered_df_cache = rows
        self.total_items = total_count

        self.update_results_list(rows)
        self._update_pagination_ui()

        # ジャンプ予約の処理
        if hasattr(self, "_pending_reveal_key") and self._pending_reveal_key:
            self._handle_pending_reveal()

    def _handle_pending_reveal(self):
        target_index = -1
        # list_item_widgets は ui_mixin で管理される
        if hasattr(self, "list_item_widgets"):
            for i, item in enumerate(self.list_item_widgets):
                if item["key"] == self._pending_reveal_key:
                    target_index = i
                    break

        if target_index != -1:
            self._set_list_cursor(target_index)
            self.focus_set()
            self.selection_info_label.configure(
                text=f"ページ {self.current_page + 1}", text_color="#28a745"
            )
            self.after(1500, self.update_selection_ui_state)

        self._pending_reveal_key = None

    # -------------------------------------------------------------------------
    # ページネーション制御 (検索状態の変更)
    # -------------------------------------------------------------------------
    def prev_page(self):
        """前のページへ移動"""
        if self.current_page > 0:
            self.current_page -= 1
            self.perform_search(reset_page=False)

    def next_page(self):
        """次のページへ移動"""
        # 現在の総件数から最大ページ数を計算
        max_page = max(0, (self.total_items - 1) // self.items_per_page)

        if self.current_page < max_page:
            self.current_page += 1
            self.perform_search(reset_page=False)

    def first_page(self):
        """最初のページへ移動"""
        if self.current_page > 0:
            self.current_page = 0
            self.perform_search(reset_page=False)

    def last_page(self):
        """最後のページへ移動"""
        max_page = max(0, (self.total_items - 1) // self.items_per_page)

        if self.current_page < max_page:
            self.current_page = max_page
            self.perform_search(reset_page=False)

    # -------------------------------------------------------------------------
    # データ取得ヘルパー (Export / Merge用)
    # -------------------------------------------------------------------------
    def _get_target_dataframe(self):
        """
        エクスポートやPDF結合のために、現在の対象データ（選択中または検索結果全体）を
        Pandas DataFrame 形式で取得する。
        """
        if not self.db_conn:
            return None, "search_results"

        target_data = []
        mode = "search_results"

        try:
            # 1. 選択されているノートがある場合 -> DBからそのノートだけを取得
            if self.selected_keys:
                mode = "selected_items"
                keys_list = list(self.selected_keys)
                placeholders = ",".join("?" * len(keys_list))

                # 必要なカラムを全て取得
                sql = f"""
                    SELECT date, time, title, pages, tags, key, memo,
                           commonplace_key, filepath, merged_pdf_filename,
                           merged_start_page, summary
                    FROM notes
                    WHERE key IN ({placeholders})
                    ORDER BY date, time
                """
                cursor = self.db_conn.cursor()
                cursor.execute(sql, keys_list)

                # 辞書リストに変換
                columns = [col[0] for col in cursor.description]
                for row in cursor.fetchall():
                    target_data.append(dict(zip(columns, row)))

            # 2. 選択がない場合 -> 現在の検索条件で【全件】取得
            else:
                # utils.fetch_notes_from_db を利用
                # (utilsのインポートが必要: from Synapsen_Nexus.utils import fetch_notes_from_db)

                # limit=-1 は SQLite で「制限なし(全件)」を意味します
                target_data = fetch_notes_from_db(
                    self.db_conn,
                    where_clause=self.current_where_clause,
                    params=list(self.current_params),
                    limit=-1,
                    offset=0,
                    sort_ascending=self.sort_ascending,
                )

            if not target_data:
                return pd.DataFrame(), mode

            # 3. DataFrameに変換
            df = pd.DataFrame(target_data)
            return df, mode

        except Exception as e:
            logger.error(f"エクスポート用データ取得エラー: {e}")
            return pd.DataFrame(), mode

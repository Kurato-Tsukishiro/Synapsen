import shutil
import customtkinter as ctk
from canvas_window import CanvasWindow
import threading
from tkinter import filedialog, messagebox
import pandas as pd
from pathlib import Path
import re
import sys
import datetime
import sqlite3
import queue
import subprocess
import platform

# 分割したモジュールをインポート
import logging

from utils import (
    load_app_config,
    open_pdf_viewer,
    build_memo_display,
    build_references_display,
    update_note_in_db,
    _update_note_links,
    delete_note_from_db,
    get_pdf_page_image,
    find_file_in_paths,
    fetch_notes_from_db,
    count_notes_from_db,
)
from search_parser import parse_query_to_sql
from list_navigator import ListNavigatorMixin

from preview_window import NotePreviewWindow
from editor_window import NoteEditorWindow
from saved_search_manager import SavedSearchManager
from graph_manager import GraphManager
from export_manager import ExportManager


# ==============================================================================
# ロギング設定の初期化
# ==============================================================================
# 親ディレクトリ(ルート)をパスに追加して logging_setup.py をインポート可能にする
current_dir = Path(__file__).parent
root_dir = current_dir.parent
if str(root_dir) not in sys.path:
    sys.path.append(str(root_dir))

try:
    from logging_setup import setup_logging

    # アプリ名を指定して初期化
    setup_logging("Synapsen_Nexus")
    logger = logging.getLogger("Nexus")  # このファイル用のロガー取得
except ImportError:
    # logging_setup.py がない場合のフォールバック（print出力）
    print("Warning: logging_setup.py not found. Logging disabled.")

    class MockLogger:
        def info(self, msg):
            print(f"[INFO] {msg}")

        def error(self, msg, exc_info=None):
            print(f"[ERROR] {msg} {exc_info if exc_info else ''}")

        def warning(self, msg):
            print(f"[WARN] {msg}")

    logger = MockLogger()
# ==============================================================================


class Synapsen_Nexus(ctk.CTk, ListNavigatorMixin):
    """
    デジタル・ツェッテルカステン風ノート管理アプリ「Synapsen Nexus」のメインアプリケーションクラス。

    目次CSVを読み込み、ノートの検索、フィルタリング、
    詳細表示、関連PDFへのアクセス機能を提供する。
    """

    def __init__(self):
        """アプリケーションを初期化し、ウィンドウと変数をセットアップする。"""
        super().__init__()
        self.icon_path = self.get_icon_path()
        self.title("Synapsen Nexus")

        self.grid_columnconfigure(0, weight=3)  # 左パネル
        self.grid_columnconfigure(1, weight=2)  # 右パネル
        self.grid_rowconfigure(1, weight=1)

        # --- ページネーション管理 ---
        self.current_page = 0
        self.items_per_page = 50
        self.total_items = 0
        self.current_where_clause = ""
        self.current_params = []
        self._pending_reveal_key = None  # ページジャンプ後のカーソル移動予約

        # --- アプリケーションの状態変数 ---
        self.pdf_root_folder = None         # 統合PDFが存在する(メイン)フォルダのルートパス
        self.pdf_archive_folder = None      # 統合PDFが存在する(アーカイブフォルダ等)サブフォルダのルートパス  # noqa: E501
        self.loaded_db_path = None          # 現在開いているDBのパス
        self.db_conn = None                 # SQLiteのDB接続オブジェクト

        self.browser_path = None            # configから読み込むブラウザパス

        self.config_data = {}               # Canvas等から参照される設定辞書
        self.paper_width = 595.276          # デフォルト A4
        self.paper_height = 841.89          # デフォルト A4

        self.key_icons = {}                 # IndexKeyごとのアイコン
        self.key_colors = {}                # IndexKeyごとの色
        self.commonplace_keys_options = []  # IndexKeyの全オプション

        self.filter_checkboxes = {}         # IndexKeyフィルターのチェックボックス変数
        self.filter_panel_expanded = False  # 左フィルターパネルが開いているか
        self.details_panel_expanded = True  # 右詳細パネルが開いているか (初期表示)

        self.sort_ascending = True          # ソート順を保持する変数 (デフォルトは (昇順/古い順))
        self.selected_keys = set()          # 選択されたノートのKeyを保持するセット

        self.setup_navigation_variables()  # Mixinの初期化メソッドを呼びだす

        self._current_search_id = 0  # 検索リクエストのID
        self._search_lock = threading.Lock()

        self._details_timer = None  # リストのクリックのタイマー

        self._preview_request_id = 0  # プレビュー生成リクエスト管理ID

        # スレッド間通信用のキュー
        self._search_result_queue = queue.Queue()

        self.filtered_df_cache = pd.DataFrame()
        self.current_selected_row = None

        # CTkImageオブジェクトへの参照を保持 (ガベージコレクション対策)
        self.preview_image_object = None

        # --- オートコンプリート関連 ---
        self.predefined_tags = []  # オートコンプリート用のタグリスト
        self.all_unique_tags = []  # 全ノート + 事前定義の統合タグリスト
        self.include_all_tags_for_autocomplete = True

        self.selected_suggestion_index = -1
        self.current_suggestions = []
        self.search_timer = None  # デバウンス（検索遅延）用タイマー
        self.suggestion_timer = None  # オートコンプリート用タイマー
        self._last_suggestion_args = None  # 予測変換の引数(query, cursor_pos, match)

        self.base_path = None  # アプリの基準パス (config.ini と同じ場所)

        # 検索マネージャのインスタンス化
        self.search_manager = SavedSearchManager(self)

        self.create_widgets()
        self.load_config()

        # ExportManagerの初期化 (config情報を渡す為 設定読み込み後)
        self.exporter = ExportManager(
            {"key_icons": self.key_icons, "key_colors": self.key_colors}
        )

        # ショートカットキーの設定
        self._setup_shortcuts()

        # ウィンドウが初めて表示されたら on_map を呼ぶ
        self.bind("<Map>", self.on_map)
        # 最大化失敗時のフォールバックサイズ指定
        self.geometry("1200x800")  # (on_mapが呼ばれる前の初期サイズ)

        # ウィンドウを閉じるときのイベントをフック
        self.protocol("WM_DELETE_WINDOW", self.on_closing)

        # 検索結果監視ループを開始
        self._poll_search_result()

    def on_map(self, event):
        """
        ウィンドウが初めて画面に描画されたときに呼び出される。
        ここで最大化を実行する。
        """
        try:
            self.unbind("<Map>")
            self.state("zoomed")
            logger.debug("ウィンドウを最大化しました。")
        except Exception as e:
            logger.error(f"ウィンドウの最大化に失敗しました: {e}")

    def on_closing(self):
        """
        アプリ終了時の処理
        1. バックアップ作成 (同日は上書き保存)
        2. 古いバックアップの削除 (日単位でのローテーション)
        """
        if self.loaded_db_path and Path(self.loaded_db_path).exists():
            try:
                db_path = Path(self.loaded_db_path)
                backup_dir = db_path.parent / "db_backups"
                backup_dir.mkdir(exist_ok=True)

                # --- 1. バックアップの作成 (同日は上書き) ---
                # ファイル名: Synapsen_Master_YYYYMMDD.db
                today_str = datetime.datetime.now().strftime("%Y%m%d")
                backup_filename = f"{db_path.stem}_{today_str}.db"
                backup_path = backup_dir / backup_filename

                # copy2 は同名ファイルがあれば上書きするため、
                # その日の最後に終了した時点のデータが残ります
                shutil.copy2(db_path, backup_path)
                logger.info(f"DBバックアップを更新しました: {backup_path.name}")

                # --- 2. 古いバックアップの削除 (ローテーション) ---
                # 保持する日数 (例: 最新7日分を残す)
                MAX_BACKUP_DAYS = 7

                # バックアップファイルを検索
                # globパターン: "Synapsen_Master_*.db"
                # 日付形式 (YYYYMMDD) は文字列ソートで時系列順になるため、sorted()だけで古い順になります
                backup_files = sorted(list(backup_dir.glob(f"{db_path.stem}_*.db")))

                # 保持数を超えている場合、古いファイルを削除
                if len(backup_files) > MAX_BACKUP_DAYS:
                    # 削除対象: リストの先頭から (総数 - 保持数) 個
                    files_to_delete = backup_files[:-MAX_BACKUP_DAYS]

                    for old_file in files_to_delete:
                        try:
                            # 念のため、今日作成したばかりのファイル(バックアップ対象)は除外
                            if old_file.name == backup_filename:
                                continue

                            old_file.unlink()
                            logger.info(
                                f"古いバックアップ(ローテーション)を削除しました: {old_file.name}"
                            )
                        except Exception as e:
                            logger.error(
                                f"バックアップ削除エラー ({old_file.name}): {e}"
                            )

            except Exception as e:
                logger.error(f"DBバックアップ処理全体でエラー: {e}")

        # DB接続を閉じる
        if self.db_conn:
            self.db_conn.close()

        self.destroy()

    def get_icon_path(self):
        """
        実行環境(.exe or .py)に応じて、
        プロジェクトルートの 'assets' フォルダにある
        'synapsen.ico' のパスを返す。
        """
        try:
            if getattr(sys, "frozen", False):
                # .exe実行の場合 (exeと同じフォルダがプロジェクトルート)
                project_root = Path(sys.executable).parent
            else:
                # .pyスクリプト実行の場合 (このファイルの親フォルダがプロジェクトルート)
                project_root = Path(__file__).parent.parent

            icon_path = project_root / "assets" / "synapsen.ico"

            if icon_path.is_file():
                return icon_path
        except Exception as e:
            logger.error(f"Error finding icon path: {e}")
        return None

    def load_config(self):
        """
        config.iniファイルからアプリケーション設定を読み込み、適用する。
        """
        try:
            # 実行ファイルのパスを基準にconfig.iniを探す
            if getattr(sys, "frozen", False):
                # .exe実行の場合 (e.g., F:\Synapsen\dist)
                # self.base_path は .exe と同じ場所
                self.base_path = Path(sys.executable).parent
            else:
                # .pyスクリプト実行の場合 (e.g., F:\Synapsen\Synapsen_Nexus)
                # self.base_path は main.py と同じ場所
                self.base_path = Path(__file__).parent

            # utilsから設定を辞書として読み込む
            config_data = load_app_config(self.base_path)

            # 読み込んだ設定をクラス属性にセット
            self.pdf_root_folder = config_data.get("pdf_root_folder", Path(""))
            self.pdf_archive_folder = config_data.get("pdf_archive_folder", None)
            self.nexus_output_folder = config_data.get(
                "nexus_output_folder", Path("Nexus_Output")
            )
            self.browser_path = config_data.get("browser_path", None)
            self.key_icons = config_data.get("key_icons", {})
            self.key_colors = config_data.get("key_colors", {})
            self.commonplace_keys_options = config_data.get(
                "commonplace_keys_options", []
            )
            self.predefined_tags = config_data.get("predefined_tags", [])

            self.include_all_tags_for_autocomplete = config_data.get(
                "include_all_tags_for_autocomplete", True
            )

            self.exclude_tags_by_default = config_data.get(
                "exclude_tags_by_default", []
            )

            if self.exclude_tags_by_default:
                # タグ名を表示して分かりやすくする (例: "除外: Archive")
                label_text = f"除外: {','.join(self.exclude_tags_by_default)}"
                if len(label_text) > 20:  # 長すぎる場合は省略
                    label_text = "除外タグ適用"

                self.exclude_tags_checkbox.configure(text=label_text)
                self.exclude_tags_checkbox.select()  # 初期状態でONにする
            else:
                # 設定がない場合はチェックボックスを無効化または非表示
                self.exclude_tags_checkbox.configure(
                    state="disabled", text="除外設定なし"
                )

            paper_size = config_data.get("paper_size", "A4").upper()
            if paper_size == "A5":
                self.paper_width = 419.528
                self.paper_height = 595.276
            else:
                self.paper_width = 595.276
                self.paper_height = 841.89

            # フィルターチェックボックスをUIに反映
            self.populate_key_filters()

            # 検索マネージャの読み込み
            try:
                # search_manager は config.ini と同じ場所(root) のパスを必要とする
                if getattr(sys, "frozen", False):
                    # .exe の場合、self.base_path (e.g., dist/) が root
                    root_path = self.base_path
                else:
                    # .py の場合、self.base_path.parent (e.g., Synapsen/) が root
                    root_path = self.base_path.parent

                self.search_manager.load_saved_searches(root_path)
            except Exception as e:
                logger.error(f"保存済み検索の読み込みエラー: {e}")

            # デフォルトDBが設定されていれば自動で読み込む
            default_db_path = config_data.get("database_path")

            if default_db_path and default_db_path.is_file():
                self.load_db_from_path(default_db_path)
            else:
                if default_db_path:
                    logger.warning(
                        f"デフォルトデータベースが見つかりません: {default_db_path}"
                    )
                # DBがない場合でも検索UIの初期化は行う
                self.perform_search()

        except FileNotFoundError as e:
            messagebox.showerror("設定エラー", str(e))
            self.destroy()
            return
        except Exception as e:
            logger.error(f"Config loading error: {e}", exc_info=True)  # ログに出す
            messagebox.showerror(
                "設定読み込みエラー", f"config.iniの読み込みに失敗しました: {e}"
            )

    # -------------------------------------------------------------------------
    # ショートカット設定メソッド
    # -------------------------------------------------------------------------
    def _setup_shortcuts(self):
        """キーボードショートカットをバインドする"""

        # --- 1. 機能ショートカット (入力中は無効化) ---

        # Rキー: ランダムノート (Random)
        self.bind("r", lambda e: self._handle_shortcut(self.show_random_note))
        self.bind("R", lambda e: self._handle_shortcut(self.show_random_note))

        # Cキー: キャンバス (Canvas)
        self.bind("c", lambda e: self._handle_shortcut(self.open_canvas))
        self.bind("C", lambda e: self._handle_shortcut(self.open_canvas))

        # Ctrl+Enter : 選択中のノートをCanvasへ送る (Send)
        self.bind(
            "<Control-Return>",
            lambda e: self._handle_shortcut(self.send_selection_to_canvas),
        )

        # Pキー: 詳細プレビュー (Preview)
        self.bind(
            "p", lambda e: self._handle_shortcut(self.open_current_note_in_preview)
        )
        self.bind(
            "P", lambda e: self._handle_shortcut(self.open_current_note_in_preview)
        )

        # Gキー: 全体グラフ (Global Graph)
        self.bind("g", lambda e: self._handle_shortcut(self.generate_and_show_graph))
        self.bind("G", lambda e: self._handle_shortcut(self.generate_and_show_graph))

        # Lキー: 関連グラフ (Local Graph)
        self.bind(
            "l", lambda e: self._handle_shortcut(self.show_local_graph_from_main_panel)
        )
        self.bind(
            "L", lambda e: self._handle_shortcut(self.show_local_graph_from_main_panel)
        )

        # Ctrl+E: 編集 (Edit)
        self.bind("<Control-e>", lambda e: self._handle_shortcut(self.open_edit_dialog))
        self.bind("<Control-E>", lambda e: self._handle_shortcut(self.open_edit_dialog))

        # F5: 再読み込み (Reload)
        self.bind("<F5>", lambda e: self._handle_shortcut(self._reload_db))

        # Ctrl+Shift+B: 左サイドバー(フィルター)の切り替え
        self.bind(
            "<Control-Shift-b>",
            lambda e: self._handle_shortcut(self.toggle_filter_panel),
        )
        self.bind(
            "<Control-Shift-B>",
            lambda e: self._handle_shortcut(self.toggle_filter_panel),
        )

        # Ctrl+B: 右サイドバー(詳細パネル)の切り替え
        self.bind(
            "<Control-b>", lambda e: self._handle_shortcut(self.toggle_details_panel)
        )
        self.bind(
            "<Control-B>", lambda e: self._handle_shortcut(self.toggle_details_panel)
        )

        # Ctrl+A: すべて選択
        self.bind("<Control-a>", lambda e: self._handle_shortcut(self.select_all_notes))
        self.bind("<Control-A>", lambda e: self._handle_shortcut(self.select_all_notes))

        # Ctrl+D: 選択解除 (Deselect)
        self.bind("<Control-d>", lambda e: self._handle_shortcut(self.clear_selection))
        self.bind("<Control-D>", lambda e: self._handle_shortcut(self.clear_selection))

        # Alt+S: ソート順切り替え
        self.bind("<Alt-s>", lambda e: self._handle_shortcut(self.toggle_sort_order))
        self.bind("<Alt-S>", lambda e: self._handle_shortcut(self.toggle_sort_order))

        # --- 2. 制御ショートカット (常時有効) ---

        # Ctrl+J: Jump to List (NEW)
        self.bind(
            "<Control-j>",
            lambda e: self._handle_shortcut(self.reveal_current_note_in_list),
        )
        self.bind(
            "<Control-J>",
            lambda e: self._handle_shortcut(self.reveal_current_note_in_list),
        )

        # Escキー: フォーカス解除
        self.bind("<Escape>", self._handle_escape)

        # Ctrl+F: 検索バーへフォーカス (Find)
        self.bind("<Control-f>", self._focus_search)
        self.bind("<Control-F>", self._focus_search)

        # --- 3. リスト操作用ショートカット ---
        self.setup_navigation_shortcuts()

    def _handle_escape(self, event):
        """
        Escキーが押されたら、メインウィンドウ自体にフォーカスを当てることで
        EntryやTextboxからフォーカスを外す。
        """
        self.focus_set()

    def _handle_shortcut(self, command):
        """
        ショートカット実行時のハンドラ。
        入力フィールド(Entry, Text)にフォーカスがある場合は無視して文字入力を優先する。
        """
        focused_widget = self.focus_get()

        # フォーカス中のウィジェットが存在し、かつ入力系クラスの場合
        if focused_widget:
            widget_class = focused_widget.winfo_class()
            # 'Entry': 1行入力 (CTkEntryの中身もこれ)
            # 'Text': 複数行入力 (CTkTextboxの中身もこれ)
            if widget_class in ["Entry", "Text"]:
                return

        # 入力中でなければコマンドを実行
        command()

    def _focus_search(self, event):
        """検索バーにフォーカスを移動し、全選択状態にする"""
        self.search_entry.focus_set()
        self.search_entry.select_range(0, "end")
        return "break"  # デフォルトの動作を抑制

    def _reload_db(self):
        """現在開いているDBを再読み込みする"""
        if self.loaded_db_path:
            self.load_db_from_path(self.loaded_db_path)

    def reveal_current_note_in_list(self):
        """
        詳細パネルに表示中のノートをリスト内で探し、カーソル移動する。
        別ページにある場合はページ遷移を行う。
        """
        # ノートが選択されていない場合は何もしない
        if self.current_selected_row is None:
            return

        target_key = self.current_selected_row.get("key")
        if not target_key:
            return

        # 1. まず現在の表示リスト内を探す
        target_index = -1
        if self.list_item_widgets:
            for i, item in enumerate(self.list_item_widgets):
                if item["key"] == target_key:
                    target_index = i
                    break

        if target_index != -1:
            # 現在のページに見つかった場合
            self._set_list_cursor(target_index)
            self.focus_set()
        else:
            # 2. 見つからない場合、別ページにあるか計算してジャンプ
            target_page = self._calculate_page_for_key(target_key)

            if target_page != -1:
                if target_page == self.current_page:
                    # 計算上は今のページにいるはずだが見つからない (フィルタで除外されている等)
                    self.selection_info_label.configure(
                        text="リスト外", text_color="orange"
                    )
                    self.after(1500, self.update_selection_ui_state)
                else:
                    # ページ遷移を実行
                    self.selection_info_label.configure(
                        text="ジャンプ中...", text_color="#E0a800"
                    )

                    self._pending_reveal_key = target_key  # 読み込み完了後の予約
                    self.current_page = target_page
                    self.perform_search(reset_page=False)  # 指定ページで再検索
            else:
                # 検索条件に合致しない (除外タグなど)
                self.selection_info_label.configure(
                    text="リスト外", text_color="orange"
                )
                self.after(1500, self.update_selection_ui_state)

    def _calculate_page_for_key(self, target_key):
        """
        指定されたKeyが、現在の検索条件・ソート順で何ページ目にあるかを計算する。
        """
        if not self.db_conn or not target_key:
            return -1

        try:
            # 現在の検索条件 (WHERE句)
            where_clause = self.current_where_clause
            params = list(self.current_params) if self.current_params else []

            # 現在のソート順
            order_dir = "ASC" if self.sort_ascending else "DESC"

            # 全件のKeyをソート順通りに取得 (軽量なので全件取得しても高速)
            sql = "SELECT key FROM notes"
            if where_clause:
                sql += f" WHERE {where_clause}"

            # utils.fetch_notes_from_db と同じソート順にする
            sql += f" ORDER BY date {order_dir}, time {order_dir}, key {order_dir}"

            cursor = self.db_conn.cursor()
            cursor.execute(sql, params)

            # 全キーを取得してインデックスを探す
            # (数万件程度ならメモリ展開しても一瞬です)
            all_keys = [row[0] for row in cursor.fetchall()]

            if target_key in all_keys:
                index = all_keys.index(target_key)
                # ページ番号 = インデックス // 1ページあたりの件数
                return index // self.items_per_page

            return -1

        except Exception as e:
            logger.error(f"ページ計算エラー: {e}")
            return -1

    # -------------------------------------------------------------------------

    def refresh_unique_tags(self):
        """
        タグリストを更新する。
        config設定(include_all_tags_for_autocomplete)によって挙動が変わる。
        """
        # 1. 事前定義タグでセットを初期化
        tags_set = set(self.predefined_tags)

        if self.include_all_tags_for_autocomplete and self.db_conn:
            try:
                # DISTINCT tags を取得して分解
                cursor = self.db_conn.cursor()
                cursor.execute("SELECT DISTINCT tags FROM notes")
                for (tags_str,) in cursor.fetchall():
                    if tags_str:
                        for t in tags_str.split(";"):
                            if t.strip():
                                tags_set.add(t.strip())
            except Exception as e:
                logger.error(f"Tags refresh error: {e}")

        self.all_unique_tags = sorted(list(tags_set))

    def create_widgets(self):
        # --- トップコンテナ (全体を包むフレーム) ---
        # row=0 に配置し、内部を2段構成にする
        top_container = ctk.CTkFrame(self)
        top_container.grid(
            row=0, column=0, columnspan=2, padx=10, pady=(10, 0), sticky="ew"
        )

        # ============================================================
        # 【1段目】 検索バー行 (DB操作 / 検索入力 / 検索保存)
        # ============================================================
        row1_frame = ctk.CTkFrame(top_container, fg_color="transparent")
        row1_frame.pack(side="top", fill="x", expand=True, pady=(5, 2), padx=5)

        # --- [左側] 基本ボタン (DB, ヘルプ) ---
        left_basic_frame = ctk.CTkFrame(row1_frame, fg_color="transparent")
        left_basic_frame.pack(side="left", padx=(0, 5))

        # "DBを開く" ボタン
        ctk.CTkButton(
            left_basic_frame, text="DB", command=self.load_database_dialog, width=50
        ).pack(side="left", padx=(0, 5))

        # 検索ヘルプボタン
        ctk.CTkButton(
            left_basic_frame, text="？", command=self.show_search_help, width=30
        ).pack(side="left", padx=0)

        # --- [右側] スマート検索 (検索保存, 呼び出し) ---
        # 検索バーより先にpackして右端を確保する
        smart_search_frame = ctk.CTkFrame(row1_frame, fg_color="transparent")
        smart_search_frame.pack(side="right", padx=(5, 0))

        self.save_search_button = ctk.CTkButton(
            smart_search_frame,
            text="検索保存",
            command=self.search_manager.save_current_search,
            width=80,
        )
        self.save_search_button.pack(side="left", padx=(0, 5))

        # 保存済み検索呼び出しボタン
        self.saved_search_combo = ctk.CTkComboBox(
            smart_search_frame,
            values=["保存済み検索..."],
            width=150,
            command=self.search_manager.on_saved_search_selected,
        )
        self.saved_search_combo.pack(side="left", padx=0)
        self.saved_search_combo.set("保存済み検索...")

        # --- [中央] 検索バー (残りのスペースを埋める) ---
        search_container = ctk.CTkFrame(row1_frame, fg_color="transparent")
        search_container.pack(side="left", fill="x", expand=True, padx=5)

        self.search_entry = ctk.CTkEntry(
            search_container,
            placeholder_text=(
                "検索 (例: (date:>=20240101 AND date:<=20240131) AND"
                + " (tag:Type_Fleeting OR tag:Question))"
            ),
        )
        self.search_entry.pack(fill="x", expand=True)

        # 検索バーのイベントバインド (既存維持)
        self.search_entry.bind("<KeyRelease>", self.handle_keyrelease)
        self.search_entry.bind("<FocusOut>", self.hide_autocomplete)
        self.search_entry.bind("<FocusIn>", self.schedule_suggestions)
        self.search_entry.bind("<Down>", self.navigate_suggestions)
        self.search_entry.bind("<Up>", self.navigate_suggestions)
        self.search_entry.bind("<Return>", self.confirm_suggestion)

        # ============================================================
        # 【2段目】 ツールバー行 (ソート / アクション / ツール)
        # ============================================================
        row2_frame = ctk.CTkFrame(top_container, fg_color="transparent")
        row2_frame.pack(side="top", fill="x", pady=(0, 5), padx=5)

        # --- [左側] 表示・フィルタリング系 ---
        view_tools_frame = ctk.CTkFrame(row2_frame, fg_color="transparent")
        view_tools_frame.pack(side="left", padx=0)

        # ソート順切り替えボタン
        self.sort_button = ctk.CTkButton(
            view_tools_frame,
            text="▲ 古い順",
            command=self.toggle_sort_order,
            width=90,
            fg_color="#585a9c",
            hover_color="#494B83",
        )
        self.sort_button.pack(side="left", padx=(0, 5))

        # 本文・メモ検索(FTS)
        self.fts_checkbox = ctk.CTkCheckBox(view_tools_frame, text="本文・メモ検索")
        self.fts_checkbox.pack(side="left", padx=(0, 10))
        self.fts_checkbox.configure(command=self._trigger_search_now)

        # 除外タグ有効化チェックボックス
        self.exclude_tags_checkbox = ctk.CTkCheckBox(view_tools_frame, text="除外タグ")
        self.exclude_tags_checkbox.pack(side="left", padx=(0, 10))
        self.exclude_tags_checkbox.configure(command=self._trigger_search_now)

        # 選択数表示ラベル
        self.selection_info_label = ctk.CTkLabel(
            view_tools_frame,
            text="選択: 0",
            font=("", 12, "bold"),
            text_color="gray",
            width=60,
        )
        self.selection_info_label.pack(side="left", padx=(0, 5))

        # 選択解除ボタン
        self.clear_selection_button = ctk.CTkButton(
            view_tools_frame,
            text="×",
            command=self.clear_selection,
            width=30,
            fg_color="#6C757D",
            hover_color="#5A6268",
        )
        self.clear_selection_button.pack(side="left", padx=0)

        # 分区切り線（視覚的な区切り）
        ctk.CTkLabel(row2_frame, text="|", text_color="gray").pack(side="left", padx=5)

        # --- [中央～右] アクション系 ---
        action_tools_frame = ctk.CTkFrame(row2_frame, fg_color="transparent")
        action_tools_frame.pack(side="left", padx=0)

        # グラフメニュー
        self.graph_menu_var = ctk.StringVar(value="グラフ表示")
        self.graph_menu = ctk.CTkOptionMenu(
            action_tools_frame,
            variable=self.graph_menu_var,
            values=["全体 (Global)", "関連 (Local)", "選択 (Selected)"],
            command=self.handle_graph_menu,
            width=130,
            fg_color="#585a9c",
            button_color="#494B83",
        )
        self.graph_menu.pack(side="left", padx=(0, 5))

        # リンクコピーボタン
        self.copy_links_button = ctk.CTkButton(
            action_tools_frame,
            text="リンクコピー",
            command=self.copy_selected_links,
            width=90,
            fg_color="#28a745",
            hover_color="#218838",
            state="disabled",
        )
        self.copy_links_button.pack(side="left", padx=(0, 5))

        # エクスポートメニュー
        self.export_menu_var = ctk.StringVar(value="エクスポート")
        self.export_menu = ctk.CTkOptionMenu(
            action_tools_frame,
            variable=self.export_menu_var,
            values=[
                "データ (CSV/TXT)",
                "統合PDF (Merge)",
                "全て (Data + PDF)",
                "MOC (Markdown)",
            ],
            command=self.handle_export_menu,
            width=130,
            fg_color="#17a2b8",
            button_color="#138496",
        )
        self.export_menu.pack(side="left", padx=(0, 5))

        # 分区切り線
        ctk.CTkLabel(row2_frame, text="|", text_color="gray").pack(side="left", padx=5)

        # --- [右側] ツール系 ---
        extra_tools_frame = ctk.CTkFrame(row2_frame, fg_color="transparent")
        extra_tools_frame.pack(side="left", padx=0)

        # ランダムノートボタン
        self.random_note_button = ctk.CTkButton(
            extra_tools_frame,
            text="閃き (R)",
            command=self.show_random_note,
            width=70,
            fg_color="#585a9c",
            hover_color="#494B83",
        )
        self.random_note_button.pack(side="left", padx=(0, 5))

        # キャンバスボタン
        self.canvas_button = ctk.CTkButton(
            extra_tools_frame,
            text="キャンバス",
            command=self.open_canvas,
            width=80,
            fg_color="#e0a800",
            hover_color="#c69500",
        )
        self.canvas_button.pack(side="left", padx=0)

        # 詳細パネルトグルボタン
        self.toggle_details_button = ctk.CTkButton(
            extra_tools_frame,
            text="▶ 詳細",
            command=self.toggle_details_panel,
            width=60,
            fg_color="transparent",
            border_width=1,
            text_color=("gray10", "gray90"),
        )
        self.toggle_details_button.pack(side="left", padx=(5, 0))

        # ============================================================
        # その他のUI要素 (初期化)
        # ============================================================

        # オートコンプリート用の非表示フレーム
        self.autocomplete_frame = ctk.CTkScrollableFrame(self, label_text="")

        # --- 左パネル (フィルタ・リスト) ---
        self.left_panel = ctk.CTkFrame(self, fg_color="transparent")
        self.left_panel.grid(row=1, column=0, padx=10, pady=10, sticky="nsew")
        self.left_panel.grid_rowconfigure(2, weight=1)
        self.left_panel.grid_columnconfigure(0, weight=1)

        # フィルターコンテナ
        filter_container = ctk.CTkFrame(self.left_panel)
        filter_container.grid(row=0, column=0, sticky="ew")
        filter_container.grid_columnconfigure(1, weight=1)
        self.toggle_filter_button = ctk.CTkButton(
            filter_container, text="", command=self.toggle_filter_panel, width=20
        )
        self.toggle_filter_button.grid(row=0, column=0, padx=5, pady=5)

        # フィルター非表示時に選択中アイコンを表示するフレーム
        self.collapsed_icons_frame = ctk.CTkFrame(
            filter_container, fg_color="transparent"
        )
        self.collapsed_icons_frame.grid(row=0, column=1, padx=5, pady=5, sticky="w")

        # IndexKeyフィルターのスクロールフレーム (初期非表示)
        self.key_filter_frame = ctk.CTkScrollableFrame(self.left_panel, label_text="")
        self.key_filter_frame.grid(row=1, column=0, padx=0, pady=(0, 5), sticky="nsew")

        # 検索結果リスト
        self.results_list = ctk.CTkScrollableFrame(
            self.left_panel, label_text="ノート一覧"
        )
        self.results_list.grid(row=2, column=0, padx=0, pady=0, sticky="nsew")

        # --- ページネーション UI ---
        self.pagination_frame = ctk.CTkFrame(
            self.left_panel, fg_color="transparent", height=40
        )
        self.pagination_frame.grid(row=3, column=0, sticky="ew", pady=(5, 0))

        self.btn_prev_page = ctk.CTkButton(
            self.pagination_frame,
            text="< 前へ",
            width=80,
            command=self.prev_page,
            state="disabled",
        )
        self.btn_prev_page.pack(side="left", padx=10)

        self.page_label = ctk.CTkLabel(self.pagination_frame, text="0 / 0")
        self.page_label.pack(side="left", expand=True)

        self.btn_next_page = ctk.CTkButton(
            self.pagination_frame,
            text="次へ >",
            width=80,
            command=self.next_page,
            state="disabled",
        )
        self.btn_next_page.pack(side="right", padx=10)

        # --- 右パネル (詳細・プレビュー) ---
        self.details_frame = ctk.CTkFrame(self)
        self.details_frame.grid(row=1, column=1, padx=(0, 10), pady=10, sticky="nsew")
        self.details_panel_visible = True  # 初期表示状態

        # グリッド設定 (上部情報(2列)、メモ、引用元、ボタン)
        self.details_frame.grid_rowconfigure(0, weight=0)  # 上部情報
        self.details_frame.grid_rowconfigure(1, weight=1)  # メモ
        self.details_frame.grid_rowconfigure(2, weight=1)  # 引用元
        self.details_frame.grid_rowconfigure(3, weight=0)  # ボタン
        self.details_frame.grid_columnconfigure(0, weight=1)

        # === 1. 上部情報エリア (2列構成) ===
        top_info_frame = ctk.CTkFrame(self.details_frame, fg_color="transparent")
        top_info_frame.grid(row=0, column=0, sticky="ew", padx=5, pady=5)
        top_info_frame.grid_columnconfigure(0, weight=1)  # テキスト情報 (伸縮)
        top_info_frame.grid_columnconfigure(1, weight=0)  # プレビュー (固定気味)

        # [左カラム] テキスト情報 (動的生成用コンテナ)
        self.text_info_frame = ctk.CTkFrame(top_info_frame, fg_color="transparent")
        self.text_info_frame.grid(row=0, column=0, sticky="nsew", padx=(0, 5))

        # [右カラム] プレビュー (親フレーム)
        self.preview_frame = ctk.CTkFrame(top_info_frame, fg_color="transparent")
        self.preview_frame.grid(row=0, column=1, sticky="n", padx=(5, 0))

        self.pdf_preview_label = ctk.CTkLabel(
            self.preview_frame,
            text="No Preview",
            fg_color="gray20",
            text_color="gray70",
            width=200,
            height=280,  # A4比率に近いサイズ固定
        )
        self.pdf_preview_label.pack()

        # === 2. メモ ===
        ctk.CTkLabel(self.details_frame, text="メモ:", anchor="w").grid(
            row=1, column=0, padx=10, pady=(5, 0), sticky="w"
        )
        self.memo_display_frame = ctk.CTkScrollableFrame(self.details_frame, height=150)
        self.memo_display_frame.grid(
            row=2, column=0, padx=10, pady=(0, 5), sticky="nsew"
        )

        # === 3. 引用元 ===
        self.references_display_frame = ctk.CTkScrollableFrame(
            self.details_frame, label_text="このノートを引用", height=100
        )
        self.references_display_frame.grid(
            row=3, column=0, padx=10, pady=5, sticky="nsew"
        )

        # === 4. ボタン ===
        self.edit_button_frame = ctk.CTkFrame(
            self.details_frame, fg_color="transparent"
        )
        self.edit_button_frame.grid(row=4, column=0, pady=10)

        # 詳細プレビューボタン
        self.open_preview_button = ctk.CTkButton(
            self.edit_button_frame,
            text="詳細プレビューで開く",
            command=self.open_current_note_in_preview,
            state="disabled",
            fg_color="#00695C",
            hover_color="#004D40",
        )
        self.open_preview_button.pack(side="left", padx=10)

        self.edit_button = ctk.CTkButton(
            self.edit_button_frame,
            text="このノートを編集",
            command=self.open_edit_dialog,
            state="disabled",
        )
        self.edit_button.pack(side="left", padx=10)

        self.delete_button = ctk.CTkButton(
            self.edit_button_frame,
            text="DBから削除",
            command=self.confirm_delete_note,
            fg_color="#D9534F",
            hover_color="#C9302C",
            state="disabled",
        )
        self.delete_button.pack(side="left", padx=10)

        # フィルターパネルの初期表示を同期
        self.sync_filter_panel_view()
        self.update_details_panel(None)  # 初期状態として「未選択」表示を行う

    def quick_search(self, query):
        """指定されたクエリを検索バーに入力して即座に検索を実行する"""
        self.search_entry.delete(0, "end")
        self.search_entry.insert(0, query)
        self.perform_search()

    def open_file_location(self, file_path_str):
        """指定されたファイルの保存場所をエクスプローラで開く"""
        path = Path(file_path_str)
        if not path.exists():
            messagebox.showerror("エラー", f"ファイルが見つかりません。\n{path}")
            return

        try:
            system_name = platform.system()
            if system_name == "Windows":
                subprocess.Popen(f'explorer /select,"{path}"')
            elif system_name == "Darwin":
                subprocess.Popen(["open", "-R", str(path)])
            else:
                subprocess.Popen(["xdg-open", str(path.parent)])
        except Exception as e:
            logger.error(f"フォルダオープンエラー: {e}")
            messagebox.showerror("エラー", f"フォルダを開けませんでした: {e}")

    def show_search_help(self):
        """検索プレフィックスのヘルプウィンドウを表示する。"""

        # 既にウィンドウが開いている場合は、それをフォーカスする
        if hasattr(self, "help_window") and self.help_window.winfo_exists():
            self.help_window.focus()
            self.help_window.grab_set()
            return

        # カスタムクラス (SearchHelpWindow) をインスタンス化する
        self.help_window = SearchHelpWindow(self)

    def toggle_sort_order(self):
        """ソート順（昇順/降順）を切り替え、検索を再実行する"""
        self.sort_ascending = not self.sort_ascending

        if self.sort_ascending:
            self.sort_button.configure(text="▲ 古い順")
        else:
            self.sort_button.configure(text="▼ 新しい順")

        # ソート順を変えたら即座に再検索してリストを更新
        self._trigger_search_now()

    # --- グラフメニューのハンドラ ---
    def handle_graph_menu(self, choice):
        """グラフメニューで選択された項目に応じて処理を分岐する"""
        if choice == "全体 (Global)":
            self.generate_and_show_graph()
        elif choice == "関連 (Local)":
            self.show_local_graph_from_main_panel()
        elif choice == "選択 (Selected)":
            self.show_selected_graph()

        # 処理後、メニューの表示を元に戻してボタンのように振る舞わせる
        self.graph_menu.set("グラフ表示")

    def handle_export_menu(self, choice):
        if choice == "データ (CSV/TXT)":
            # (情報)エクスポートメソッドを呼び出す
            self.export_search_results(include_pdf=False)
        elif choice == "統合PDF (Merge)":
            # PDF結合メソッドを呼び出す
            self.merge_and_export_pdf()
        elif choice == "全て (Data + PDF)":
            # 情報及び統合PDF出力を両方呼び出す
            self.export_search_results(include_pdf=True)
        elif choice == "MOC (Markdown)":
            self.create_moc_markdown()

        self.export_menu.set("エクスポート")

    # --- オートコンプリート関連メソッド ---
    def handle_keyrelease(self, event):
        """検索バーでのキー入力（リリース）イベントを処理する。"""

        # 1. 無視すべきキー（ナビゲーション、修飾キー、ファンクションキー等）
        ignored_keys = (
            "Return",
            "Escape",
            "Left",
            "Up",
            "Down",
            "Right",
            "Home",
            "End",
            "Prior",
            "Next",
            "Control_L",
            "Control_R",
            "Shift_L",
            "Shift_R",
            "Alt_L",
            "Alt_R",
            "Caps_Lock",
            "Tab",
            "ISO_Left_Tab",
            "F1",
            "F2",
            "F3",
            "F4",
            "F5",
            "F6",
            "F7",
            "F8",
            "F9",
            "F10",
            "F11",
            "F12",
            # 半角/全角
            "Zenkaku_Hankaku",
            "Kanji",
            # 無変換、変換
            "Muhenkan",
            "Henkan",
            # ひらがな、カタカナ
            "Hiragana",
            "Katakana",
            "Hiragana_Katakana",
            # 英数
            "Eisu_toggle",
            "Alphanumeric",
        )

        if event.keysym in ignored_keys:
            return

        # 2. ショートカット操作（Ctrl+F, Ctrl+Aなど）を無視する
        # event.state のビットマスクで Ctrl(4) または Alt(8 or 131072) が押されているか判定
        is_ctrl = (event.state & 0x0004) != 0
        is_alt = (event.state & 0x0008) != 0 or (event.state & 0x20000) != 0

        if is_ctrl or is_alt:
            return

        # 上記以外のキー (文字入力、Delete, BackSpaceなど) の場合のみ
        # 予測変換と検索をスケジュールする
        self.schedule_suggestions()
        self.schedule_search()

    def update_suggestions(self, query, cursor_pos, match_value):
        """
        'tag:' プレフィックス入力中にオートコンプリート候補を更新する。
        """
        self.selected_suggestion_index = -1
        # 'tag:abc' の 'abc' の部分 (group 2) を取得
        last_tag_word = match_value.group(2).strip() if match_value else ""
        target_list = self.all_unique_tags or self.predefined_tags

        if last_tag_word == "":
            # 'tag:' 直後は全リスト
            suggestions = target_list
        else:
            # 前方一致検索# 'tag:Py' のように入力中の場合は、前方一致検索
            last_word_lower = last_tag_word.lower()
            suggestions = [
                tag for tag in target_list if tag.lower().startswith(last_word_lower)
            ]

        # 上下キーでの移動 (navigate_suggestions) のために引数を保存
        self._last_suggestion_args = (query, cursor_pos, match_value)

        if suggestions:
            # show_autocomplete にも引数を渡す
            self.show_autocomplete(suggestions, query, cursor_pos, match_value)
        else:
            self.hide_autocomplete()

    def show_autocomplete(self, suggestions, query, cursor_pos, match_value):
        """オートコンプリートの候補リストウィンドウを表示する。"""
        self.current_suggestions = suggestions
        for widget in self.autocomplete_frame.winfo_children():
            widget.destroy()

        for i, suggestion in enumerate(suggestions):
            # 選択中のインデックスに基づいてハイライト色を決定
            fg_color = (
                "gray30" if i == self.selected_suggestion_index else "transparent"
            )

            btn = ctk.CTkButton(
                self.autocomplete_frame,
                text=suggestion,
                fg_color=fg_color,
                text_color=ctk.ThemeManager.theme["CTkLabel"]["text_color"],
                anchor="w",
                # select_suggestion に必要な情報をラムダで渡す
                command=lambda s=suggestion: self.select_suggestion(
                    s, query, cursor_pos, match_value
                ),
            )
            btn.pack(fill="x", padx=5, pady=2)

        # 検索バーの真下に配置 (元のコード)
        x = self.search_entry.winfo_rootx() - self.winfo_rootx()
        y = (
            self.search_entry.winfo_rooty()
            - self.winfo_rooty()
            + self.search_entry.winfo_height()
        )
        width = self.search_entry.winfo_width()
        height = min(200, len(suggestions) * 35)

        self.autocomplete_frame.configure(width=width, height=height)
        self.autocomplete_frame.place(x=x, y=y)
        self.autocomplete_frame.lift()

    def select_suggestion(self, suggestion, query, cursor_pos, match_value):
        """オートコンプリート候補をクリックまたはEnterで選択したときの処理。"""

        prefix_part = ""
        suffix_part = query[cursor_pos:]  # カーソルより後ろのテキスト

        if match_value:
            # 'tag:Py' のように入力中の場合
            # (group 2 が 'Py' の部分)
            # 'tag:' の直前までを取得
            prefix_part = query[: match_value.start(2)]
        else:
            # 'tag:' と入力した直後の場合 (match_value は None)
            # カーソル位置までをそのまま使用
            prefix_part = query[:cursor_pos]

            # 'tag:' の直後にスペースがない場合、自動でスペースを追加
            if not prefix_part.endswith(" "):
                suggestion = " " + suggestion

        # クエリを再構築
        new_query = f"{prefix_part}{suggestion} {suffix_part}"

        # 新しいカーソル位置 = 'tag:' + 'Python' + ' ' の直後
        new_cursor_pos = len(prefix_part) + len(suggestion) + 1

        # UIに反映
        self.search_entry.delete(0, "end")
        self.search_entry.insert(0, new_query)
        self.search_entry.focus_force()
        self.search_entry.icursor(new_cursor_pos)  # カーソル位置を更新

        self.hide_autocomplete()
        self._trigger_search_now()  # 検索を即時実行

    def hide_autocomplete(self, event=None):
        """オートコンプリートウィンドウを非表示にする。"""
        # 少し遅延させて非表示にし、クリックイベントが発火できるようにする
        self.after(200, lambda: self.autocomplete_frame.place_forget())

    def navigate_suggestions(self, event):
        """キーボードの上下矢印キーで候補リストを移動する。"""
        if (
            not self.autocomplete_frame.winfo_ismapped()
            or not self.current_suggestions
            or self._last_suggestion_args is None
        ):
            return

        num = len(self.current_suggestions)
        if event.keysym == "Down":
            self.selected_suggestion_index = (self.selected_suggestion_index + 1) % num
        elif event.keysym == "Up":
            self.selected_suggestion_index = (
                self.selected_suggestion_index - 1 + num
            ) % num

        # 選択項目がリストに表示されるようにスクロール
        self.autocomplete_frame._parent_canvas.yview_moveto(
            self.selected_suggestion_index / num
        )

        # 選択ハイライトを更新 (保存しておいた引数を使う)
        q, c, m = self._last_suggestion_args
        self.show_autocomplete(self.current_suggestions, q, c, m)
        return "break"  # 他のキーバインドを抑制

    def confirm_suggestion(self, event):
        """Enterキーで選択中の候補を確定する。"""
        if (
            self.autocomplete_frame.winfo_ismapped()
            and self.selected_suggestion_index != -1
            and self._last_suggestion_args is not None
        ):

            # 保存しておいた引数を取得
            q, c, m = self._last_suggestion_args
            # select_suggestion を呼び出す
            self.select_suggestion(
                self.current_suggestions[self.selected_suggestion_index], q, c, m
            )
            return "break"  # 検索が二重に実行されるのを防ぐ

        # 候補が選択されていない場合は、通常の検索を実行
        self._trigger_search_now()
        self.hide_autocomplete()

    # --- フィルターパネル関連メソッド ---
    def sync_filter_panel_view(self):
        """フィルターパネルの開閉状態をUIに同期させる。"""
        if self.filter_panel_expanded:
            self.key_filter_frame.grid()
            self.toggle_filter_button.configure(text="▼ IndexKey フィルター")
        else:
            self.key_filter_frame.grid_remove()
            self.toggle_filter_button.configure(text="▶ IndexKey フィルター")
        self.update_collapsed_filter_view()

    def toggle_filter_panel(self):
        """フィルターパネルの開閉状態を切り替える。"""
        self.filter_panel_expanded = not self.filter_panel_expanded
        self.sync_filter_panel_view()

    def update_collapsed_filter_view(self):
        """
        フィルターパネルが閉じているときに、
        選択中のフィルターアイコンを表示する。
        """
        for widget in self.collapsed_icons_frame.winfo_children():
            widget.destroy()

        if not self.filter_panel_expanded:
            if not self.filter_panel_expanded:
                # 選択されている IndexKey をリストアップ
                selected_keys = []
                for key, var in self.filter_checkboxes.items():
                    if var.get() == "1":
                        selected_keys.append(key)

            if not selected_keys:
                ctk.CTkLabel(self.collapsed_icons_frame, text="", font=("", 16)).pack(
                    side="left"
                )
            else:
                for key in selected_keys:
                    icon = self.key_icons.get(key.lower(), "•")
                    color = self.key_colors.get(key.lower(), "gray")
                    ctk.CTkLabel(
                        self.collapsed_icons_frame,
                        text=icon,
                        text_color=color,
                        font=("", 16),
                    ).pack(side="left", padx=2)

    # --- 右サイドバー（詳細パネル）のトグル ---
    def toggle_details_panel(self):
        """右側の詳細パネルの表示/非表示を切り替える"""
        if self.details_panel_expanded:
            self.details_frame.grid_remove()
            self.details_panel_expanded = False
            self.toggle_details_button.configure(text="◀ 詳細")
        else:
            self.details_frame.grid()
            self.details_panel_expanded = True
            self.toggle_details_button.configure(text="▶ 詳細")

    # --- データ読み込み・検索実行メソッド ---
    def load_database_dialog(self):
        """「目次データベースを開く」ボタンの動作。ファイルダイアログを開く。"""
        filepath = filedialog.askopenfilename(
            title="目次データベースファイルを選択",
            filetypes=[("SQLite Database", "*.db"), ("All files", "*.*")],
        )
        if not filepath:
            return
        self.load_db_from_path(Path(filepath))

    def load_db_from_path(self, filepath: Path, key_to_redisplay: str = None):
        """
        指定されたパスからDBを読み込み、DataFrameを更新する。
        utils.load_sql_data_file を使用する。

        Args:
            filepath (Path): 読み込むDBファイルのパス。
            key_to_redisplay (str, optional):
                読み込み後に詳細ペインに再表示するノートのキー。
                Noneの場合は詳細ペインをクリアする。
        """
        try:
            # --- 1. 読み取り専用接続 (db_conn) のクローズ/再生成 ---
            if self.db_conn:
                self.db_conn.close()

            self.loaded_db_path = filepath
            # SQLite接続 (読み取り専用)
            self.db_conn = sqlite3.connect(f"file:{filepath}?mode=ro", uri=True)

            # --- テーブル構造チェック ---
            conn_write = None
            try:
                conn_write = sqlite3.connect(filepath)  # 書き込みモードで接続
                cursor = conn_write.cursor()

                # リンクテーブルの作成
                cursor.execute(
                    """
                CREATE TABLE IF NOT EXISTS note_links (
                    source_key TEXT NOT NULL,
                    target_key TEXT NOT NULL,
                    PRIMARY KEY (source_key, target_key)
                )
                """
                )
                # 引用元検索(target_key)を高速化するインデックス
                cursor.execute(
                    """
                CREATE INDEX IF NOT EXISTS idx_target_key
                    ON note_links (target_key)
                """
                )
                conn_write.commit()
                logger.info("'note_links' テーブルの存在を確認・作成しました。")

            except Exception as e_tbl:
                logger.error(f"'note_links' テーブルの作成に失敗: {e_tbl}")
                if conn_write:
                    conn_write.rollback()
            finally:
                if conn_write:
                    conn_write.close()

            # UIリセット & 初期検索
            self.current_page = 0
            self.refresh_unique_tags()  # タグリスト更新はSQLで行うよう要修正(後述)
            self.perform_search()  # ここで最初の50件がロードされる

            # 再表示処理
            if key_to_redisplay:
                # 特定のノートだけ個別に取得して表示
                notes = fetch_notes_from_db(self.db_conn, "key = ?", [key_to_redisplay])
                if notes:
                    self.show_details(
                        pd.Series(notes[0])
                    )  # 互換性のためSeriesに変換して渡す
                else:
                    self.clear_details()
            else:
                self.clear_details()

        except Exception as e:
            messagebox.showerror("DBエラー", str(e))

    def populate_key_filters(self):
        for widget in self.key_filter_frame.winfo_children():
            widget.destroy()
        self.filter_checkboxes.clear()

        for key in self.commonplace_keys_options:
            var = ctk.StringVar(value="0")
            row_frame = ctk.CTkFrame(self.key_filter_frame, fg_color="transparent")
            row_frame.pack(anchor="w", padx=10, pady=2, fill="x")

            icon = self.key_icons.get(key.lower(), "•")
            color = self.key_colors.get(key.lower(), "gray")

            ctk.CTkLabel(
                row_frame, text=icon, text_color=color, font=("", 16), width=20
            ).pack(side="left")

            ctk.CTkCheckBox(
                row_frame,
                text=key,
                variable=var,
                onvalue="1",
                offvalue="0",
                command=self._trigger_search_now,  # チェック時に検索を再実行
            ).pack(side="left", expand=True, fill="x")

            self.filter_checkboxes[key] = var

    def _poll_search_result(self):
        """
        バックグラウンドスレッドからの検索結果を監視し、
        キューにデータがあればメインスレッドで更新処理を行う。
        """
        try:
            while True:
                # キューから (search_id, result_data) を取り出す
                # result_data は (rows, total) のタプル
                search_id, result_data = self._search_result_queue.get_nowait()
                self._on_search_complete(search_id, result_data)
        except queue.Empty:
            pass
        finally:
            self.after(100, self._poll_search_result)

    def perform_search(self, reset_page=True):
        """
        検索処理のエントリーポイント。
        UIスレッドで直接検索せず、バックグラウンドスレッドを開始します。
        """
        if not self.db_conn:
            return

        if reset_page:
            self.current_page = 0

        user_query = self.search_entry.get().strip()
        include_full_text = self.fts_checkbox.get()

        # --- 検索条件の構築 (メインスレッド) ---

        # 1. 除外タグ
        exclusion_query = ""
        if self.exclude_tags_checkbox.get() == 1 and self.exclude_tags_by_default:
            parts = [f"-tag:{t}" for t in self.exclude_tags_by_default]
            exclusion_query = " ".join(parts)

        full_query_str = (
            f"({user_query}) AND ({exclusion_query})"
            if user_query and exclusion_query
            else (user_query or exclusion_query)
        )

        # 2. IndexKey フィルター
        selected_filters = [
            k for k, v in self.filter_checkboxes.items() if v.get() == "1"
        ]

        # 3. SQLへのパース
        where_parts = []
        params = []

        # クエリ文字列 -> SQL
        if full_query_str:
            q_sql, q_params = parse_query_to_sql(full_query_str, include_full_text)
            if q_sql:
                where_parts.append(q_sql)
                params.extend(q_params)

        # IndexKey -> SQL
        if selected_filters:
            placeholders = ",".join(["?"] * len(selected_filters))
            where_parts.append(f"commonplace_key IN ({placeholders})")
            params.extend(selected_filters)

        final_where = " AND ".join(where_parts) if where_parts else ""

        # 状態保存 (ページネーション用)
        self.current_where_clause = final_where
        self.current_params = params

        # ID更新
        with self._search_lock:
            self._current_search_id += 1
            search_id = self._current_search_id

        # ローディング表示
        self.results_list.configure(label_text="検索中...")
        self.configure(cursor="watch")

        # スレッド開始
        thread = threading.Thread(
            target=self._execute_search_worker,
            args=(
                search_id,
                self.loaded_db_path,  # パスを渡してスレッド内で接続
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
        try:
            conn = sqlite3.connect(f"file:{db_path}?mode=ro", uri=True)

            # 1. 総件数取得
            total = count_notes_from_db(conn, where, params)

            # 2. データ取得
            offset = page * limit
            rows = fetch_notes_from_db(conn, where, params, limit, offset, asc)

            conn.close()

            self._search_result_queue.put((search_id, (rows, total)))

        except Exception as e:
            logger.error(f"Search thread error: {e}")
            # エラー時も形式を合わせる
            self._search_result_queue.put((search_id, ([], 0)))

    def _on_search_complete(self, search_id, result_data):
        """検索完了時のコールバック"""
        if isinstance(result_data, tuple):
            rows, total_count = result_data
        else:
            return

        self.configure(cursor="")

        with self._search_lock:
            if search_id != self._current_search_id:
                return

        self.filtered_df_cache = rows
        self.total_items = total_count

        self.update_results_list(rows)
        self._update_pagination_ui()

        # 【追加】 ジャンプ予約がある場合の処理
        if self._pending_reveal_key:
            target_index = -1
            for i, item in enumerate(self.list_item_widgets):
                if item["key"] == self._pending_reveal_key:
                    target_index = i
                    break

            if target_index != -1:
                self._set_list_cursor(target_index)
                self.focus_set()  # フォーカスをリストに戻す
                # 成功メッセージ
                self.selection_info_label.configure(
                    text=f"ページ {self.current_page + 1}", text_color="#28a745"
                )
                self.after(1500, self.update_selection_ui_state)

            # 予約をクリア
            self._pending_reveal_key = None

    def _update_pagination_ui(self):
        max_page = max(0, (self.total_items - 1) // self.items_per_page)

        self.page_label.configure(text=f"{self.current_page + 1} / {max_page + 1}")

        if self.current_page > 0:
            self.btn_prev_page.configure(state="normal")
        else:
            self.btn_prev_page.configure(state="disabled")

        if self.current_page < max_page:
            self.btn_next_page.configure(state="normal")
        else:
            self.btn_next_page.configure(state="disabled")

    def next_page(self):
        self.current_page += 1
        self.perform_search(reset_page=False)

    def prev_page(self):
        if self.current_page > 0:
            self.current_page -= 1
            self.perform_search(reset_page=False)

    def _trigger_search_now(self):
        """
        デバウンスタイマーをキャンセルし、検索を即座に実行する。
        Enterキー押下時やチェックボックス変更時に使用する。
        """
        # 検索のタイマー
        if self.search_timer:
            self.after_cancel(self.search_timer)
            self.search_timer = None

        # 予測変換のタイマー
        if self.suggestion_timer:
            self.after_cancel(self.suggestion_timer)
            self.suggestion_timer = None

        # 本体の検索を実行
        self.perform_search()

    def schedule_search(self):
        """
        検索の実行を遅延させる（デバウンス）。
        既にあるタイマーをキャンセルし、新しいタイマーを設定する。
        """
        # 既存のタイマーがあればキャンセル
        if self.search_timer:
            self.after_cancel(self.search_timer)

        # 待機時間を 650ミリ秒 (0.65秒) に設定
        self.search_timer = self.after(650, self.perform_search)

    def schedule_suggestions(self, event=None):
        """
        オートコンプリートの更新を遅延させる（デバウンス）。
        [変更] 'tag:' プレフィックス入力中のみ予測変換を実行する。
        """
        if self.suggestion_timer:
            self.after_cancel(self.suggestion_timer)
            self.suggestion_timer = None

        query = self.search_entry.get()

        # 現在のカーソル位置までのクエリを取得
        cursor_pos = self.search_entry.index(ctk.INSERT)
        query_to_cursor = query[:cursor_pos]

        # 正規表現: 'tag:' (または 'tags:') の後に文字を入力中か
        # (大文字小文字を無視, ^|\s|\( は行頭/空白/括弧の意)
        tag_value_pattern = r"(?i)(^|\s|\()tags?:([^\s\)]*)$"
        match_value = re.search(tag_value_pattern, query_to_cursor)

        # 正規表現: 'tag:' (または 'tags:') を入力した直後か
        tag_prefix_pattern = r"(?i)(^|\s|\()tags?:\s*$"
        match_prefix = re.search(tag_prefix_pattern, query_to_cursor)

        if match_prefix or match_value:
            # 'tag:' が入力されている場合のみ、予測変換を実行 (200ms後)
            # update_suggestions に必要な情報を渡す
            self.suggestion_timer = self.after(
                200,
                lambda q=query, c=cursor_pos, m=match_value: self.update_suggestions(
                    q, c, m
                ),
            )
        else:
            # 'tag:' 以外が入力されている場合は、予測変換を即座に隠す
            self.hide_autocomplete()
            self._last_suggestion_args = None  # 引数キャッシュをクリア

    def show_random_note(self):
        """
        「閃き」ボタン押下時。
        現在の検索条件（または全データ）からランダムに1件のノートを取得し表示する。
        """
        if not self.db_conn:
            return

        try:
            # 現在の検索条件 (WHERE句とパラメータ) を利用
            where_clause = self.current_where_clause
            params = list(self.current_params) if self.current_params else []

            # ランダムに1件取得するSQL
            sql = "SELECT * FROM notes"
            if where_clause:
                sql += f" WHERE {where_clause}"

            sql += " ORDER BY RANDOM() LIMIT 1"

            cursor = self.db_conn.cursor()
            cursor.execute(sql, params)
            row = cursor.fetchone()

            if row:
                # カラム名付きの辞書に変換
                columns = [col[0] for col in cursor.description]
                note_data = dict(zip(columns, row))

                # 詳細表示
                self.show_details(note_data)

                # リスト内にあればカーソル移動
                target_key = note_data.get("key")
                target_index = -1
                if self.list_item_widgets:
                    for i, item in enumerate(self.list_item_widgets):
                        if item["key"] == target_key:
                            target_index = i
                            break

                if target_index != -1:
                    self._set_list_cursor(target_index)
            else:
                messagebox.showinfo("情報", "表示できるノートがありません。")

        except Exception as e:
            logger.error(f"ランダムノート表示エラー: {e}")
            messagebox.showerror("エラー", f"失敗しました: {e}")

    def open_canvas(self, background=False, notes_to_add=None):
        """キャンバスウィンドウを開く"""
        if hasattr(self, "canvas_window") and self.canvas_window.winfo_exists():
            # 既に開いている場合
            if not background:
                if self.canvas_window.state() == "iconic":
                    self.canvas_window.deiconify()

                self.canvas_window.lift()
                self.canvas_window.focus_force()

            # 追加リクエストがあれば実行
            if notes_to_add:
                self.canvas_window.add_selected_notes(target_keys=notes_to_add)
            return

        # 新規作成
        self.canvas_window = CanvasWindow(
            self, background=background, initial_notes=notes_to_add
        )

    def send_selection_to_canvas(self):
        """
        現在選択中のノートをCanvasに追加する。
        Canvasが開いていない場合は自動で開く。
        フォーカスはNexusに残す（連続作業用）。
        """
        if not self.selected_keys:
            # 選択がない場合、ラベルを一瞬赤くして通知
            self.selection_info_label.configure(text="選択なし!", text_color="red")
            self.after(1000, self.update_selection_ui_state)  # 1秒後に元に戻す
            return

        # 1. Canvasをバックグラウンドモードで開く (またはノートを渡す)
        self.open_canvas(background=True, notes_to_add=self.selected_keys)

        # 2. 成功メッセージ
        count = len(self.selected_keys)
        self.selection_info_label.configure(
            text=f"Canvasに追加: {count}件", text_color="#28a745"
        )
        self.after(1500, self.update_selection_ui_state)

        # 3. フォーカスをNexusに維持
        self.focus_force()

    # --- UI更新・表示メソッド ---
    def update_results_list(self, rows):
        """
        検索結果リストUIを更新する。
        """
        # 既存ウィジェットの削除
        for widget in self.results_list.winfo_children():
            widget.destroy()

        self.list_item_widgets = []  # 参照リストをリセット
        self.list_cursor_index = -1  # カーソルリセット
        self.list_anchor_index = -1  # アンカーのリセット

        # ラベル表示の更新
        if hasattr(self, "total_items") and self.total_items > 0:
            page_str = (
                f" (ページ {self.current_page + 1})"
                if hasattr(self, "current_page")
                else ""
            )
            label = f"検索結果: {self.total_items} 件{page_str}"
        else:
            label = f"検索結果 ({len(rows)}件)"

        self.results_list.configure(label_text=label)

        for row in rows:
            # 行フレームの作成
            item_frame = ctk.CTkFrame(self.results_list, fg_color="transparent")
            item_frame.pack(fill="x", padx=5, pady=2)

            # --- チェックボックス ---
            note_key = row.get("key")
            is_selected = note_key in self.selected_keys

            # チェックボックスの状態変数
            chk_var = ctk.StringVar(value="on" if is_selected else "off")

            def on_toggle(k=note_key, v=chk_var):
                self.toggle_note_selection(k, v)

            checkbox = ctk.CTkCheckBox(
                item_frame,
                text="",
                width=24,
                variable=chk_var,
                onvalue="on",
                offvalue="off",
                command=on_toggle,
            )
            if not note_key:
                checkbox.configure(state="disabled")

            checkbox.pack(side="left", padx=(0, 5))

            cp_key = str(row.get("commonplace_key", "")).lower()
            icon = self.key_icons.get(cp_key, "•")
            color = self.key_colors.get(cp_key, "gray")

            icon_label = ctk.CTkLabel(
                item_frame, text=icon, text_color=color, font=("", 16), width=20
            )
            icon_label.pack(side="left")

            display_text = f"[{row.get('date')}] {row.get('title', 'N/A')}"
            text_label = ctk.CTkLabel(item_frame, text=display_text, anchor="w")
            text_label.pack(side="left", fill="x", expand=True)

            # --- ウィジェットリストへの保存 ---
            # 現在のインデックスを保存（クリック時のカーソル更新用）
            current_widget_index = len(self.list_item_widgets)

            # リストに追加 (フレーム本体, データ行, チェックボックス変数)
            self.list_item_widgets.append(
                {"frame": item_frame, "data": row, "chk_var": chk_var, "key": note_key}
            )

            # --- イベントバインド ---
            def create_show_details_handler(note_row=row, idx=current_widget_index):
                def handler(event):
                    # 1. 見た目のカーソル移動は即座に行う (レスポンス確保)
                    self._set_list_cursor(idx)

                    # 2. 連続クリックされた場合、前の予約をキャンセル
                    if self._details_timer:
                        self.after_cancel(self._details_timer)

                    # 3. 0.25秒後に詳細表示を実行
                    # (この待ち時間の間にダブルクリックがあれば、そちらが優先される)
                    self._details_timer = self.after(
                        250, lambda: self.show_details(note_row)
                    )

                return handler

            def create_open_pdf_handler(note_row=row, idx=current_widget_index):
                def handler(event):
                    self._set_list_cursor(idx)
                    self.open_pdf(note_row)

                return handler

            show_details_command = create_show_details_handler()
            # item_frame自体へのクリックは詳細表示
            item_frame.bind("<Button-1>", show_details_command)
            # アイコンやテキストへのクリックも詳細表示
            icon_label.bind("<Button-1>", show_details_command)
            text_label.bind("<Button-1>", show_details_command)

            # ダブルクリックでPDFを開く
            open_pdf_command = create_open_pdf_handler()
            item_frame.bind("<Double-Button-1>", open_pdf_command)
            icon_label.bind("<Double-Button-1>", open_pdf_command)
            text_label.bind("<Double-Button-1>", open_pdf_command)

        # リスト更新後、スクロール位置を最上部にリセットする
        self.results_list._parent_canvas.yview_moveto(0)

    def toggle_note_selection(self, key, var):
        """チェックボックスの切り替え時の処理"""
        if var.get() == "on":
            self.selected_keys.add(key)
        else:
            self.selected_keys.discard(key)

        self.update_selection_ui_state()

    def select_all_notes(self):
        """現在のページ（表示されているノート）をすべて選択状態にする"""
        # データがない場合は何もしない
        if not self.filtered_df_cache:
            return

        # 表示中のノートのKeyを取得
        keys_in_view = [
            item["key"] for item in self.filtered_df_cache if item.get("key")
        ]

        # 選択セットに追加
        self.selected_keys.update(keys_in_view)

        # UI更新
        self.update_selection_ui_state()

        # リストのチェックボックス表示を更新するために再描画
        # (件数が多いと一瞬ラグがあるかもしれませんが、整合性を保つため必要です)
        self.update_results_list(self.filtered_df_cache)

    def clear_selection(self):
        """選択をすべて解除する"""
        self.selected_keys.clear()
        self.update_selection_ui_state()
        # リストのチェックボックス表示を更新するため、現在の検索結果でリストを再描画
        if self.filtered_df_cache is not None:
            self.update_results_list(self.filtered_df_cache)
        self.clear_details()

    def update_selection_ui_state(self):
        """選択数に応じてラベル表示とボタン状態を更新"""
        count = len(self.selected_keys)
        self.selection_info_label.configure(text=f"選択: {count}")

        # リンクコピーボタンの制御
        if count > 0:
            # 選択がある時は強調色（黄色/オレンジ系）にする
            self.selection_info_label.configure(text_color="#E0a800")
            self.copy_links_button.configure(state="normal")
            # 選択グラフも有効化されていることを視覚的に示すため、メニューは常に有効のまま
        else:
            # 0件の時はグレーアウト
            self.selection_info_label.configure(text_color="gray")
            self.copy_links_button.configure(state="disabled")

    def copy_selected_links(self):
        """
        選択されたノートのリンク文字列（[[Key: Title]]）を生成し、
        クリップボードにコピーする。
        """
        if not self.selected_keys or not self.db_conn:
            return

        try:
            # 選択キーのリスト化
            target_keys = list(self.selected_keys)
            if not target_keys:
                return

            # SQLでタイトルを取得
            placeholders = ",".join("?" * len(target_keys))
            sql = (
                "SELECT key, title FROM notes "
                f"WHERE key IN ({placeholders}) ORDER BY date, time"
            )

            cursor = self.db_conn.cursor()
            cursor.execute(sql, target_keys)
            rows = cursor.fetchall()

            link_texts = []
            for r in rows:
                key, title = r[0], r[1]
                link_texts.append(f"[[{key}: {title}]]")

            if not link_texts:
                return

            # 改行区切りで結合
            clipboard_text = "\n".join(link_texts)

            # クリップボードへコピー
            self.clipboard_clear()
            self.clipboard_append(clipboard_text)
            self.update()

            messagebox.showinfo(
                "コピー完了",
                f"{len(link_texts)}件のリンクをクリップボードにコピーしました。",
                parent=self,
            )

        except Exception as e:
            logger.error(f"リンクコピーエラー: {e}")
            messagebox.showerror("エラー", f"リンクのコピーに失敗しました: {e}")

    def show_selected_graph(self):
        """
        選択されたノート(self.selected_keys)のみでグラフを表示する。
        """
        if not self.selected_keys:
            messagebox.showinfo("情報", "ノートが選択されていません。")
            return

        if not self.db_conn:
            return

        try:
            keys_list = list(self.selected_keys)
            placeholders = ",".join("?" * len(keys_list))

            # GraphManagerに必要なカラムを取得
            sql = f"SELECT * FROM notes WHERE key IN ({placeholders})"

            # Pandasを使ってSQL結果を直接DataFrameにする
            selected_df = pd.read_sql_query(sql, self.db_conn, params=keys_list)

            if selected_df.empty:
                messagebox.showinfo("情報", "データが見つかりません。")
                return

            logger.info(f"Selected Graph: {len(selected_df)} notes")
            self.generate_and_show_graph(target_df=selected_df)

        except Exception as e:
            logger.error(f"選択グラフ生成エラー: {e}")
            messagebox.showerror("エラー", f"グラフ生成に失敗しました: {e}")

    def clear_details(self):
        self.update_details_panel(None)

    def append_link_to_selected_notes(self, key_to_link, title_to_link):
        """
        メイン画面で選択中の全ノートに対し、
        指定されたKeyとTitleのリンクをメモ欄の末尾に追記する。
        """
        selected_keys_to_update = self.selected_keys

        if not selected_keys_to_update:
            messagebox.showwarning(
                "リンク失敗",
                "リンクの追記先となるノートがメイン画面で選択されていません。",
                parent=self,
            )
            return

        if not self.loaded_db_path:
            messagebox.showerror(
                "DBエラー", "データベースが読み込まれていません。", parent=self
            )
            return

        # リンク文字列を生成
        link_text_to_append = f"[[{key_to_link}: {title_to_link}]]"

        conn = None
        updated_keys = []
        try:
            # 1. 書き込み用のDB接続を開始
            conn = sqlite3.connect(self.loaded_db_path)
            cursor = conn.cursor()

            for key in selected_keys_to_update:
                # 2. 現在のメモを取得
                cursor.execute("SELECT memo FROM notes WHERE key = ?", (key,))
                memo_data = cursor.fetchone()
                current_memo = memo_data[0] if memo_data and memo_data[0] else ""

                # 3. 既にリンクが含まれていないかチェック
                if link_text_to_append in current_memo:
                    continue

                # 4. メモを更新 (末尾に改行＋リンクを追加)
                new_memo = current_memo.strip() + f"\n{link_text_to_append}\n"

                # 5. DBを更新 (notes テーブル)
                cursor.execute(
                    "UPDATE notes SET memo = ? WHERE key = ?", (new_memo, key)
                )

                # 6. リンクテーブルも更新 (utils._update_note_links を使用)
                _update_note_links(cursor, key, new_memo)

                updated_keys.append(key)

            # 7. トランザクションをコミット
            conn.commit()

            if updated_keys:
                messagebox.showinfo(
                    "リンク完了",
                    f"{len(updated_keys)}件のノートにリンクを追記しました。\n\n"
                    f"(対象Key: {', '.join(updated_keys)})",
                    parent=self,
                )

                # メイン画面の詳細表示が更新対象に含まれていたら、再描画
                if (
                    self.current_selected_row is not None
                    and self.current_selected_row.get("key") in updated_keys
                ):
                    self.show_details(self.current_selected_row)
            else:
                messagebox.showinfo(
                    "リンク不要",
                    "選択中のノートには、既にリンクが追加されています。",
                    parent=self,
                )

        except Exception as e:
            if conn:
                conn.rollback()
            messagebox.showerror(
                "リンクエラー", f"リンクの追記に失敗しました:\n{e}", parent=self
            )
        finally:
            if conn:
                conn.close()

    def open_current_note_in_preview(self):
        """
        メインペインの「詳細プレビューで開く」ボタンから呼び出される。
        'full' (拡大) モードでプレビューを開く。
        """
        if self.current_selected_row is None:
            messagebox.showwarning(
                "ノート未選択", "詳細プレビューを開くノートが選択されていません。"
            )
            return

        key_to_open = self.current_selected_row.get("key")
        if key_to_open:
            self.open_preview_window(key_to_open, default_view_mode="full")

    def open_preview_window(self, key, default_view_mode="compact", ui_master=None):
        """
        指定されたキーのノートを新しい「簡易プレビュー」ウィンドウで開く。
        【修正】pandas依存を排除し、DBからデータを取得する。
        """
        from utils import fetch_notes_from_db  # ローカルインポート (念のため)

        if not self.db_conn:
            messagebox.showwarning("データなし", "データベースが読み込まれていません。")
            return

        # 1. DBからメタデータを取得 (リストが返る)
        notes = fetch_notes_from_db(self.db_conn, "key = ?", [key])

        if not notes:
            messagebox.showwarning(
                "ノート不明", f"ID '{key}' に一致するノートが見つかりませんでした。"
            )
            return

        # 辞書として取得
        note_data = notes[0].copy()

        # 2. memo, full_text, summary をDBから確実に取得
        try:
            cursor = self.db_conn.cursor()
            cursor.execute(
                "SELECT memo, full_text, summary FROM notes WHERE key = ?", (key,)
            )
            db_data = cursor.fetchone()
            if db_data:
                note_data["memo"] = str(db_data[0]) if db_data[0] is not None else ""
                note_data["full_text"] = (
                    str(db_data[1]) if db_data[1] is not None else ""
                )
                note_data["summary"] = str(db_data[2]) if db_data[2] is not None else ""
        except Exception as e:
            logger.error(f"プレビュー用のDBデータ取得エラー: {e}")

        # 3. プレビューウィンドウのインスタンスを作成
        preview_win = NotePreviewWindow(
            self, note_data, default_view_mode, ui_master=ui_master
        )
        preview_win.focus()

    def show_details(self, row_data):
        """
        選択されたノートの詳細を右ペインに表示する。
        (実処理は update_details_panel に委譲)

        Args:
            row_data (pd.Series): 表示するノートの行データ。
        """
        # pd.Series に加えて dict も許可する
        # (sqlite3.Row は dict(row) で変換済み想定ですが、念のため許可リストに入れるか、dict変換する)
        if not isinstance(row_data, (pd.Series, dict)):
            # row_data が sqlite3.Row の場合も考慮して dict 化を試みる
            try:
                row_data = dict(row_data)
            except (ValueError, TypeError):
                logger.error(
                    f"Error: show_details に不正なデータ型が渡されました: {type(row_data)}"
                )
                self.clear_details()
                return

        self.update_details_panel(row_data)

    def _update_preview_ui(self, request_id, pil_image):
        """
        別スレッドで生成されたプレビュー画像をUIに反映するコールバック
        """
        # 最新のリクエストでなければ無視（連打時の古い画像などは破棄）
        if request_id != self._preview_request_id:
            return

        # 既存のプレビューエリアをクリア
        for widget in self.preview_frame.winfo_children():
            widget.destroy()

        if pil_image:
            try:
                new_image = ctk.CTkImage(
                    light_image=pil_image,
                    dark_image=pil_image,
                    size=(pil_image.width, pil_image.height),
                )
                self.preview_image_object = new_image  # 参照保持
                self.pdf_preview_label = ctk.CTkLabel(
                    self.preview_frame, text="", image=new_image, fg_color="transparent"
                )
                self.pdf_preview_label.pack()
            except Exception as e:
                logger.error(f"プレビュー画像設定エラー: {e}")
                self._show_preview_error("Display Error")
        else:
            # 取得失敗 (zlib error等含む)
            self._show_preview_error("No Preview")

    def _show_preview_error(self, message):
        """プレビュー失敗時のエラー表示"""
        self.pdf_preview_label = ctk.CTkLabel(
            self.preview_frame,
            text=message,
            fg_color="gray20",
            text_color="gray70",
            width=200,
            height=280,
        )
        self.pdf_preview_label.pack()
        self.preview_image_object = None

    def update_details_panel(self, row_data):
        """
        右パネル（詳細情報）の内容を更新する。
        【修正】
          1. pandas依存を排除し、SQLと辞書リストを使用。
          2. PDFプレビュー生成を非同期(スレッド)で実行し、UIフリーズを回避。
        """
        current_data = None
        backlinks_data = []  # 辞書リストとして初期化

        # --- 1. データの前処理 (DB取得) ---
        if row_data is not None:
            # 辞書としてコピー
            current_data = dict(row_data).copy()
            current_key = current_data.get("key", "")

            if self.db_conn and current_key:
                try:
                    cursor = self.db_conn.cursor()

                    # (A) 最新のメモ・概要を取得
                    cursor.execute(
                        "SELECT memo, summary FROM notes WHERE key = ?", (current_key,)
                    )
                    data = cursor.fetchone()
                    if data:
                        current_data["memo"] = (
                            str(data[0]) if data[0] is not None else ""
                        )
                        current_data["summary"] = (
                            str(data[1]) if data[1] is not None else ""
                        )

                    # (B) 引用元 (Backlinks) をSQLで取得
                    sql_backlinks = """
                        SELECT n.key, n.title, n.date, n.commonplace_key
                        FROM notes n
                        JOIN note_links l ON n.key = l.source_key
                        WHERE l.target_key = ?
                        ORDER BY n.date DESC, n.time DESC
                    """
                    cursor.execute(sql_backlinks, (current_key,))
                    rows = cursor.fetchall()

                    # 辞書リストに変換
                    for r in rows:
                        backlinks_data.append(
                            {
                                "key": r[0],
                                "title": r[1],
                                "date": r[2],
                                "commonplace_key": r[3],
                            }
                        )

                except Exception as e:
                    logger.error(f"詳細データ取得エラー: {e}")

            # 選択状態を更新
            self.current_selected_row = current_data

            # ボタン有効化
            self.edit_button.configure(state="normal")
            self.delete_button.configure(state="normal")
            self.open_preview_button.configure(state="normal")

        else:
            # クリア時は選択状態もクリア
            self.current_selected_row = None
            self.edit_button.configure(state="disabled")
            self.delete_button.configure(state="disabled")
            self.open_preview_button.configure(state="disabled")

        # --- 2. UI構築 ---
        # テキスト情報の再構築
        for widget in self.text_info_frame.winfo_children():
            widget.destroy()

        self.text_info_frame.grid_columnconfigure(1, weight=1)
        current_row = 0

        # ヘルパー: 行追加
        def add_row(label, value_widget_func):
            nonlocal current_row
            ctk.CTkLabel(
                self.text_info_frame, text=label, anchor="w", text_color="gray"
            ).grid(row=current_row, column=0, sticky="nw", pady=2)
            value_widget_func(current_row)
            current_row += 1

        if current_data is None:
            # 未選択状態の描画
            for label in [
                "タイトル:",
                "キー:",
                "Index Key:",
                "ファイル:",
                "タグ:",
                "概要:",
            ]:
                add_row(
                    label,
                    lambda r: ctk.CTkLabel(
                        self.text_info_frame, text="-", anchor="w"
                    ).grid(row=r, column=1, sticky="ew"),
                )

            # プレビューリセット
            for widget in self.preview_frame.winfo_children():
                widget.destroy()
            self.pdf_preview_label = ctk.CTkLabel(
                self.preview_frame,
                text="No Preview",
                fg_color="gray20",
                text_color="gray70",
                width=200,
                height=280,
            )
            self.pdf_preview_label.pack()
            self.preview_image_object = None

            # メモ・引用元クリア
            for widget in self.memo_display_frame.winfo_children():
                widget.destroy()
            for widget in self.references_display_frame.winfo_children():
                widget.destroy()
            self.references_display_frame.configure(label_text="このノートを引用 (0件)")
            return

        # --- 選択状態の描画 ---

        # 1. タイトル
        add_row(
            "タイトル:",
            lambda r: ctk.CTkLabel(
                self.text_info_frame,
                text=current_data.get("title", ""),
                wraplength=390,
                justify="left",
                anchor="w",
            ).grid(row=r, column=1, sticky="ew"),
        )

        # 2. キー
        add_row(
            "キー:",
            lambda r: ctk.CTkLabel(
                self.text_info_frame, text=current_data.get("key", ""), anchor="w"
            ).grid(row=r, column=1, sticky="ew"),
        )

        # 3. Index Key
        cp_key = current_data.get("commonplace_key", "")

        def create_ikey_widget(r):
            if cp_key:
                f = ctk.CTkFrame(self.text_info_frame, fg_color="transparent")
                f.grid(row=r, column=1, sticky="ew")
                icon = self.key_icons.get(cp_key.lower(), "•")
                col = self.key_colors.get(cp_key.lower(), "gray")
                ctk.CTkLabel(
                    f, text=icon, text_color=col, font=("", 16), width=20
                ).pack(side="left")
                ctk.CTkButton(
                    f,
                    text=cp_key,
                    fg_color="transparent",
                    text_color=("#1F6AA5", "#63B8FF"),
                    hover_color=("gray85", "gray25"),
                    anchor="w",
                    height=20,
                    command=lambda k=cp_key: self.quick_search(f"ikey:{k}"),
                ).pack(side="left", fill="x", expand=True)
            else:
                ctk.CTkLabel(self.text_info_frame, text="-", anchor="w").grid(
                    row=r, column=1, sticky="ew"
                )

        add_row("Index Key:", create_ikey_widget)

        # 4. ファイル (ここを修正)
        merged_name = current_data.get("merged_pdf_filename", "")
        file_path_str = current_data.get("filepath", "")
        display_filename = (
            merged_name
            if (merged_name and merged_name != "nan")
            else (Path(file_path_str).name if file_path_str else "(不明)")
        )

        # 開くべきファイルのパス検索
        target_open_path = file_path_str
        if merged_name and merged_name != "nan":
            search_paths = []
            if self.pdf_root_folder:
                search_paths.append(Path(self.pdf_root_folder))
            if self.pdf_archive_folder:
                search_paths.append(Path(self.pdf_archive_folder))
            if self.loaded_db_path:
                search_paths.append(Path(self.loaded_db_path).parent)

            found = find_file_in_paths(merged_name, search_paths)
            if found:
                target_open_path = str(found)

        def create_file_widget(r):
            f = ctk.CTkFrame(self.text_info_frame, fg_color="transparent")
            f.grid(row=r, column=1, sticky="ew")
            if target_open_path:
                ctk.CTkButton(
                    f,
                    text="📂",
                    width=30,
                    height=24,
                    fg_color="gray70",
                    hover_color="gray50",
                    text_color="black",
                    command=lambda p=target_open_path: self.open_file_location(p),
                ).pack(side="left", padx=(0, 5))
            if display_filename != "(不明)":
                ctk.CTkButton(
                    f,
                    text=display_filename,
                    fg_color="transparent",
                    text_color=("#1F6AA5", "#63B8FF"),
                    hover_color=("gray85", "gray25"),
                    anchor="w",
                    height=20,
                    command=lambda fn=display_filename: self.quick_search(f"file:{fn}"),
                ).pack(side="left", fill="x", expand=True)
            else:
                ctk.CTkLabel(f, text=display_filename, anchor="w").pack(side="left")

        add_row("ファイル:", create_file_widget)

        # 5. タグ
        tags_list = [t for t in str(current_data.get("tags", "")).split(";") if t]

        def create_tags_widget(r):
            if tags_list:
                s = ctk.CTkScrollableFrame(
                    self.text_info_frame,
                    height=35,
                    orientation="horizontal",
                    fg_color="transparent",
                )
                s.grid(row=r, column=1, sticky="ew")
                for t in tags_list:
                    ctk.CTkButton(
                        s,
                        text=t,
                        font=("", 11),
                        fg_color=("gray80", "gray30"),
                        text_color=("black", "white"),
                        hover_color=("gray70", "gray40"),
                        height=20,
                        width=30,
                        command=lambda tag=t: self.quick_search(f"tag:{tag}"),
                    ).pack(side="left", padx=2)
            else:
                ctk.CTkLabel(self.text_info_frame, text="-", anchor="w").grid(
                    row=r, column=1, sticky="ew"
                )

        add_row("タグ:", create_tags_widget)

        # 6. 概要
        add_row(
            "概要:",
            lambda r: ctk.CTkLabel(
                self.text_info_frame,
                text=current_data.get("summary", ""),
                wraplength=390,
                justify="left",
                anchor="w",
            ).grid(row=r, column=1, sticky="ew"),
        )

        # --- ▼ 3. PDFプレビューの表示 ▼ ---

        # 1. リクエストIDを更新 (古いスレッドの結果を破棄するため)
        self._preview_request_id += 1
        req_id = self._preview_request_id

        # 2. まず「Loading...」を表示しておく
        for widget in self.preview_frame.winfo_children():
            widget.destroy()

        self.pdf_preview_label = ctk.CTkLabel(
            self.preview_frame,
            text="Loading...",
            fg_color="gray20",
            text_color="gray70",
            width=200,
            height=280,
        )
        self.pdf_preview_label.pack()

        # 3. 画像生成を別スレッドで実行する関数
        def run_preview_generation():
            image = None
            try:
                # get_pdf_page_image は Series を期待する場合があるため念のため変換
                # (utils.py の実装によっては dict でも動くが、安全策)
                try:
                    data_for_utils = pd.Series(current_data)
                except Exception:
                    data_for_utils = current_data

                image = get_pdf_page_image(
                    data_for_utils,
                    self.loaded_db_path,
                    self.pdf_root_folder,
                    max_width=200,
                    pdf_archive_folder=self.pdf_archive_folder,
                )
            except Exception as e:
                # ここでのエラーはログに出して無視（メイン処理を止めない）
                logger.warning(f"プレビュー生成失敗(非同期): {e}")

            # 4. メインスレッドでUI更新を予約 (スレッドセーフな呼び出し)
            self.after(0, lambda: self._update_preview_ui(req_id, image))

        # スレッド開始
        threading.Thread(target=run_preview_generation, daemon=True).start()

        # --- 4. メモと引用元の表示 ---
        for widget in self.memo_display_frame.winfo_children():
            widget.destroy()

        # DB接続を渡す (SQL化対応)
        build_memo_display(
            self.memo_display_frame,
            current_data.get("memo", ""),
            self.db_conn,
            lambda k: self.open_preview_window(k, default_view_mode="compact"),
            450,
        )

        # 辞書リストを渡す (SQL化対応)
        build_references_display(
            self.references_display_frame,
            backlinks_data,
            lambda k: self.open_preview_window(k, default_view_mode="compact"),
            self.key_icons,
            self.key_colors,
        )

    # --- PDF関連メソッド ---
    def jump_to_key(self, key):
        """
        (現在未使用) 指定されたキーを検索窓に入力し、検索する。

        Args:
            key (str): 検索するノートの 'key' (ID)。
        """
        self.search_entry.delete(0, "end")
        self.search_entry.insert(0, f"key:{key}")
        self.perform_search()

    def jump_to_pdf(self, key):
        """
        (現在未使用) 指定されたキーを持つノートのPDFを開く。
        """
        if not self.db_conn:
            return

        try:
            cursor = self.db_conn.cursor()
            cursor.execute("SELECT * FROM notes WHERE key = ?", (key,))
            row = cursor.fetchone()

            if not row:
                messagebox.showwarning(
                    "ノート不明", f"ID '{key}' が見つかりませんでした。"
                )
                return

            columns = [col[0] for col in cursor.description]
            note_data = dict(zip(columns, row))

            self.open_pdf(note_data)

        except Exception as e:
            logger.error(f"jump_to_pdf error: {e}")

    def open_pdf(self, row_data):
        """
        ノートデータに基づきPDFを開くラッパーメソッド。
        utils.open_pdf_viewer を呼び出す。

        Args:
            row_data (pd.Series): PDFを開く対象のノートデータ。
        """
        # utilsの関数に、必要な設定値（パス情報）と共に渡す
        open_pdf_viewer(
            row_data,
            self.loaded_db_path,
            self.pdf_root_folder,
            self.browser_path,
            pdf_archive_folder=self.pdf_archive_folder,
        )

    # --- グラフ生成メソッド ---
    def generate_and_show_graph(self, output_path=None, target_df=None):
        """
        グラフを生成して表示する。

        Args:
            output_path (Path, optional): 保存先パス。
            target_df (pd.DataFrame, optional):
                グラフ化対象のデータフレーム。
                Noneの場合は現在の検索結果(self.filtered_df_cache)を使用する。
        """
        df = None

        # 1. データソースの決定
        if target_df is not None:
            df = target_df
        else:
            # 選択なし -> 検索結果全体 (ただしグラフなので上限を設ける)
            limit = 500
            if self.db_conn:
                try:
                    from utils import fetch_notes_from_db

                    rows = fetch_notes_from_db(
                        self.db_conn,
                        where_clause=self.current_where_clause,
                        params=list(self.current_params),
                        limit=limit,
                        offset=0,
                        sort_ascending=self.sort_ascending,
                    )

                    if not rows:
                        if not output_path:
                            messagebox.showinfo(
                                "グラフ表示",
                                "グラフ化するノートがありません。",
                                parent=self,
                            )
                        return

                    df = pd.DataFrame(rows)

                    # 件数が上限に達していたら通知
                    # (fetch_notes_from_db は limit 件数分しか返さないため、len == limit なら上限到達とみなす)
                    if len(df) >= limit and not output_path:
                        messagebox.showwarning(
                            "グラフ表示",
                            f"件数が多いため、検索結果の上位 {limit} 件のみを表示します。\n"
                            "全件表示したい場合はフィルタリングを行ってください。",
                            parent=self,
                        )
                except Exception as e:
                    logger.error(f"グラフ用データ取得エラー: {e}")
                    if not output_path:
                        messagebox.showerror(
                            "エラー", f"データ取得に失敗しました: {e}", parent=self
                        )
                    return

        # 2. DataFrame空チェック
        if df is None or df.empty:
            if not output_path:
                messagebox.showinfo("グラフ表示", "データがありません。", parent=self)
            return

        try:
            # 3. GraphManager に処理を委譲
            generated_path = GraphManager.generate_graph_html(
                df,
                self.key_icons,
                self.key_colors,
                self.loaded_db_path,
                self.pdf_root_folder,
                output_path=output_path,
                db_conn=self.db_conn,
                pdf_archive_folder=self.pdf_archive_folder,
            )

            # 4. ブラウザで表示 (ファイル保存モードでない場合)
            if not output_path:
                GraphManager.open_graph(generated_path)

        except Exception as e:
            logger.error(f"Graph error: {e}", exc_info=True)
            if not output_path:
                messagebox.showerror("エラー", f"グラフ生成失敗: {e}", parent=self)

    def show_local_graph_from_main_panel(self):
        """
        メインパネルの「関連グラフ」メニューから呼ばれるラッパー。
        """
        if self.current_selected_row is None:
            messagebox.showinfo(
                "情報", "ローカルグラフを表示するノートを選択してください。"
            )
            return

        center_key = self.current_selected_row.get("key")
        if center_key:
            self.show_local_graph(center_key)

    def show_local_graph(self, center_key: str):
        """
        中心ノートと、それに関連する（リンク/被リンク）ノートでグラフを表示する。
        """
        if not center_key or not self.db_conn:
            return

        related_keys = {center_key}

        try:
            cursor = self.db_conn.cursor()

            # A. このノートが引用している先 (Forward Links)
            cursor.execute(
                "SELECT target_key FROM note_links WHERE source_key = ?", (center_key,)
            )
            for row in cursor.fetchall():
                related_keys.add(row[0])

            # B. このノートを引用している元 (Backlinks)
            cursor.execute(
                "SELECT source_key FROM note_links WHERE target_key = ?", (center_key,)
            )
            for row in cursor.fetchall():
                related_keys.add(row[0])

            # C. データ取得 (DataFrame化)
            keys_list = list(related_keys)
            if not keys_list:
                return

            placeholders = ",".join("?" * len(keys_list))
            sql = f"SELECT * FROM notes WHERE key IN ({placeholders})"

            local_df = pd.read_sql_query(sql, self.db_conn, params=keys_list)

            if local_df.empty:
                messagebox.showinfo("情報", "関連ノートが見つかりませんでした。")
                return

            logger.info(f"Local Graph: {len(local_df)} notes")
            self.generate_and_show_graph(target_df=local_df)

        except Exception as e:
            logger.error(f"ローカルグラフ生成エラー: {e}")
            messagebox.showerror("エラー", f"失敗しました: {e}")

    # --- エクスポート用データ取得ヘルパー ---
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
                from utils import fetch_notes_from_db

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

    # --- エクスポート機能 (修正) ---
    def export_search_results(self, include_pdf=False):
        """
        検索結果(または選択中)のデータをエクスポートする。
        """
        # 【修正】ヘルパーを使ってDataFrameを取得
        target_df, mode = self._get_target_dataframe()

        if target_df is None or target_df.empty:
            messagebox.showinfo("エクスポート", "対象データがありません。", parent=self)
            return

        initial_dir = None
        if hasattr(self, "nexus_output_folder") and self.nexus_output_folder:
            if self.nexus_output_folder.exists():
                initial_dir = str(self.nexus_output_folder)

        export_folder = filedialog.askdirectory(
            title="エクスポート先を選択",
            initialdir=initial_dir,
        )
        if not export_folder:
            return

        try:
            # ExportManager に処理を委譲
            success, result_path = self.exporter.execute_export(
                target_df=target_df,
                query_text=self.search_entry.get(),
                export_folder_path=export_folder,
                mode=mode,
                include_pdf=include_pdf,
                loaded_db_path=self.loaded_db_path,
                pdf_root_folder=self.pdf_root_folder,
                pdf_archive_folder=self.pdf_archive_folder,
            )

            msg = f"エクスポート完了:\n{result_path}"
            messagebox.showinfo("完了", msg, parent=self)

        except Exception as e:
            logger.error(f"エクスポート実行エラー: {e}", exc_info=True)
            messagebox.showerror(
                "エクスポートエラー", f"失敗しました:\n{e}", parent=self
            )

    def merge_and_export_pdf(self):
        """メニューから「統合PDF」単体を選んだ場合のラッパー"""
        # 対象データの決定
        target_df, _ = self._get_target_dataframe()

        if target_df is None or target_df.empty:
            messagebox.showinfo("PDF結合", "対象データがありません。", parent=self)
            return

        time_text = datetime.datetime.now().strftime("%Y%m%d")

        initial_dir = None
        if hasattr(self, "nexus_output_folder") and self.nexus_output_folder:
            if self.nexus_output_folder.exists():
                initial_dir = str(self.nexus_output_folder)

        save_path = filedialog.asksaveasfilename(
            title="統合PDFを保存",
            defaultextension=".pdf",
            filetypes=[("PDF Files", "*.pdf")],
            initialfile=f"Merged_Notes_{time_text}.pdf",
            initialdir=initial_dir,
            parent=self,
        )
        if not save_path:
            return

        try:
            if self.exporter.merge_pdf(
                target_df,
                Path(save_path),
                self.pdf_root_folder,
                self.loaded_db_path,  # DBパスも渡す必要がある場合があります
                pdf_archive_folder=self.pdf_archive_folder,
            ):
                messagebox.showinfo(
                    "完了", f"PDFを保存しました:\n{save_path}", parent=self
                )
            else:
                messagebox.showwarning(
                    "失敗", "結合可能なPDFが見つかりませんでした。", parent=self
                )
        except Exception as e:
            logger.error(f"PDF結合エラー: {e}")
            messagebox.showerror("エラー", f"PDF結合中にエラーが発生しました:\n{e}")

    def create_moc_markdown(self):
        """MOC (Markdown) 生成処理のラッパー"""
        # 対象データの決定 (選択中 > 検索結果)
        target_df, _ = self._get_target_dataframe()

        if target_df is None or target_df.empty:
            messagebox.showinfo("MOC作成", "対象データがありません。", parent=self)
            return

        # 保存先ダイアログ
        time_text = datetime.datetime.now().strftime("%Y%m%d")
        initial_dir = None
        if hasattr(self, "nexus_output_folder") and self.nexus_output_folder:
            if self.nexus_output_folder.exists():
                initial_dir = str(self.nexus_output_folder)
        save_path = filedialog.asksaveasfilename(
            title="MOC (Markdown) を保存",
            defaultextension=".md",
            filetypes=[("Markdown Files", "*.md")],
            initialfile=f"MOC_{time_text}.md",
            initialdir=initial_dir,
        )
        if not save_path:
            return

        # 生成実行
        try:
            if self.exporter.generate_moc_markdown(
                target_df, str(Path(save_path)), str(self.loaded_db_path)
            ):
                messagebox.showinfo(
                    "完了",
                    f"MOC (Markdown) ファイルを保存しました:\n{save_path}\n\n"
                    "このファイルを Normalisierer で処理し PDF 化すると、\n"
                    "ノートへのファイルパスと Index Key の装飾が反映された PDF を作成できます。",
                    parent=self,
                )
            else:
                messagebox.showerror(
                    "失敗", "MOCファイルの生成に失敗しました。", parent=self
                )
        except Exception as e:
            logger.error(f"MOC生成エラー: {e}")
            messagebox.showerror("エラー", f"MOC生成中にエラーが発生しました:\n{e}")

    # --- DB編集・削除メソッド ---
    def open_edit_dialog(self, note_data=None):
        """
        編集ボタン押下時、またはリンククリック時に編集ウィンドウを開く。
        """
        if note_data is None:
            if self.current_selected_row is None:
                messagebox.showerror("エラー", "編集するノートが選択されていません。")
                return
            note_data = self.current_selected_row

        if not self.db_conn:
            messagebox.showwarning("データなし", "データベースが読み込まれていません。")
            return

        # 辞書化 (pd.Series や sqlite3.Row 対策)
        if not isinstance(note_data, dict):
            try:
                note_data = dict(note_data)
            except (ValueError, TypeError):
                pass

        key_to_edit = note_data.get("key")
        if not key_to_edit:
            messagebox.showerror("エラー", "Keyが不明です。")
            return

        # 最新の memo, summary をDBから取得
        note_data_with_memo = note_data.copy()
        try:
            cursor = self.db_conn.cursor()
            cursor.execute(
                "SELECT memo, summary FROM notes WHERE key = ?", (key_to_edit,)
            )
            row = cursor.fetchone()
            if row:
                note_data_with_memo["memo"] = str(row[0]) if row[0] is not None else ""
                note_data_with_memo["summary"] = (
                    str(row[1]) if row[1] is not None else ""
                )
            else:
                note_data_with_memo["memo"] = ""
                note_data_with_memo["summary"] = ""

        except Exception as e:
            logger.error(f"編集用データ取得エラー: {e}")
            note_data_with_memo["memo"] = ""  # 失敗時は空メモ

        # 編集ウィンドウ (書き込み可能) のインスタンスを作成
        editor_win = NoteEditorWindow(
            self,
            note_data_with_memo,  # 最新メモ入りのデータを渡す
            self.commonplace_keys_options,
            self.all_unique_tags,  # 全タグを渡す
            self.save_edit_callback,
        )
        editor_win.focus()

    def save_edit_callback(self, new_data_dict):
        """
        編集ウィンドウ(NoteEditorWindow)の保存ボタンから呼び出される。

        Args:
            new_data_dict (dict): 編集された新しいノートデータ。
        """
        if not self.loaded_db_path:
            raise Exception("データベースのパスが不明です。")

        key_to_update = new_data_dict.get("key")
        if not key_to_update:
            raise Exception("更新対象のKeyが不明です。")

        conn = None
        try:
            # 1. 書き込み用のDB接続を開始
            conn = sqlite3.connect(self.loaded_db_path)

            # 2. utilsのDB更新関数を呼び出す (connを渡す)
            update_note_in_db(conn, key_to_update, new_data_dict)

            # 3. トランザクションをコミット
            conn.commit()

            # 4. 変更をUIに反映するため、DBを再読み込み
            self.load_db_from_path(self.loaded_db_path, key_to_redisplay=key_to_update)

        except Exception as e:
            if conn:
                conn.rollback()
            messagebox.showerror(
                "保存エラー", f"データベースの更新に失敗しました:\n{e}", parent=self
            )
        finally:
            if conn:
                conn.close()

    def confirm_delete_note(self):
        """
        削除ボタン押下時。確認ダイアログを表示する。
        """
        if self.current_selected_row is None:
            messagebox.showerror("エラー", "削除するノートが選択されていません。")
            return

        key_to_delete = self.current_selected_row.get("key")
        title_to_delete = self.current_selected_row.get("title")

        if not key_to_delete:
            messagebox.showerror("エラー", "Keyが不明なため削除できません。")
            return

        # 最終確認
        answer = messagebox.askyesno(
            "削除の最終確認",
            "以下のノートをマスターデータベースから完全に削除しますか？\n\n"
            f"Key: {key_to_delete}\n"
            f"Title: {title_to_delete}\n\n"
            "この操作は元に戻せません。",
            parent=self,
        )

        if answer:
            conn = None
            try:
                # 1. 書き込み用のDB接続を開始
                conn = sqlite3.connect(self.loaded_db_path)

                # 2. utilsのDB削除関数を呼び出す (connを渡す)
                delete_note_from_db(conn, key_to_delete)

                # 3. トランザクションをコミット
                conn.commit()

                # 4. 変更をUIに反映するため、DBを再読み込み
                logger.info(
                    f"ノート {key_to_delete} を削除しました。DBを再読み込みします。"
                )
                self.load_db_from_path(self.loaded_db_path)

                messagebox.showinfo(
                    "削除完了", f"ノート {key_to_delete} を削除しました。", parent=self
                )

            except Exception as e:
                if conn:
                    conn.rollback()
                messagebox.showerror(
                    "削除エラー",
                    f"データベースからの削除に失敗しました:\n{e}",
                    parent=self,
                )
            finally:
                if conn:
                    conn.close()


# ==============================================================================
# 検索ヘルプウィンドウ
# ==============================================================================
class SearchHelpWindow(ctk.CTkToplevel):
    """
    検索ヘルプ専用のToplevelウィンドウ。
    メインアイコンを強制的に継承するため、iconbitmapメソッドをオーバーライドする。
    """

    def __init__(self, parent_app):
        super().__init__(parent_app)
        self._custom_icon_path = None  # 強制設定するアイコンパス

        if hasattr(parent_app, "icon_path") and parent_app.icon_path:
            self._custom_icon_path = str(parent_app.icon_path)
            # --- 初期アイコンをすぐに設定 ---
            if self._custom_icon_path:
                try:
                    # 親クラス(Toplevel)の iconbitmap を直接呼び出す
                    super().iconbitmap(self._custom_icon_path)
                except Exception as e:
                    logger.error(f"Initial icon set error: {e}")

        self.title("ヘルプ (検索・ショートカット)")
        self.geometry("600x700")
        self.transient(parent_app)
        self.grab_set()

        help_text = """
----------------------------------------------------------------------------------------------------
■ アプリケーション ショートカット一覧
----------------------------------------------------------------------------------------------------

[リスト操作 (キーボード)]
↑ / ↓ : リスト内を移動
Home / End : 先頭 / 末尾へ移動
PageUp / PageDn : ページ送り
Space : 選択切り替え (チェックボックス ON/OFF)
Enter : 詳細を表示 (右ペイン更新)
Shift + Enter : ノートのPDFを開く
Shift + 移動 : 範囲選択 (標準)
Ctrl + Shift + 移動 : 範囲選択 (追加モード)

[リスト操作 (コマンド)]
Ctrl + A : すべて選択
Ctrl + D : 選択解除
Ctrl + J : リストへ移動
Alt + S  : ソート順切り替え (昇順/降順)
Ctrl + E : 選択中のノートを編集
Ctrl + Enter : 選択中のノートをCanvasへ送る

[画面表示・遷移]
R : 閃き (ランダムノート表示)
C : キャンバスを開く
P : 詳細プレビューを開く
G : 全体グラフ (Global) を表示
L : 関連グラフ (Local) を表示
F5 : データベース再読み込み
Ctrl + B : 右パネル（詳細）の開閉
Ctrl + Shift + B : 左パネル（フィルター）の開閉

[検索・入力]
Ctrl + F : 検索バーへフォーカス (全選択)
Esc : 入力欄からフォーカスを外す

----------------------------------------------------------------------------------------------------
■ Synapsen Nexus 検索クエリ リファレンス
----------------------------------------------------------------------------------------------------

■ 基本
- 検索語をスペースで区切ると `AND` 検索になります。
  (例: `Type_Permanent 薬物動態学`)
- `OR` を使用すると `OR` 検索ができます。
  (例: `Type_Fleeting OR Question`)
- `()` でグループ化できます。
  (例: `(tag:Type_Fleeting OR tag:Question) AND (ikey:学習 OR ikey:タスク)`)
- `-` (ハイフン) を検索語の前に付けると `NOT` 検索になります。
  (例: `衛生 -memo:古い`)

----------------------------------------------------------------------------------------------------
■ プレフィックスとエイリアス

`title: (キーワード)`
- ノートのタイトルを検索します (部分一致)。

`key: (ID)`
- ノートのユニークIDを**部分一致**で検索します。
  (例: `key:202401` や `key:0930` など)

`[[Key]]` (エイリアス)
- ノートのユニークIDを**完全一致**で検索します。
  (例: `[[20240101090000: タイトル]]` は `20240101090000` のみヒット)

`tag: (キーワード)` (エイリアス: `tags:`)
- タグを検索します (部分一致)。入力補完 (`tag:`) が利用可能です。

`ikey: (キーワード)` (エイリアス: `cpkey:`, `indexkey:`)
- Index Key (コモンプレイスキー) を検索します (部分一致)。

`memo: (キーワード)`
- メモ欄を検索します (部分一致・「本文・メモ検索」OFFでも強制検索)。

`fulltext: (キーワード)` (エイリアス: `text:`)
- PDF本文を検索します (部分一致・「本文・メモ検索」OFFでも強制検索)。

`fulltext: (キーワード)` (エイリアス: `text:`)
- PDF本文を検索します (部分一致・「本文・メモ検索」OFFでも強制検索)。

`filename: (キーワード)` (エイリアス: `file:`)
- 統合PDFのファイル名を検索します (部分一致)。
  (例: `filename:202410` → 2024年10月の統合PDFに含まれるノートを抽出)

----------------------------------------------------------------------------------------------------
■ 特殊なプレフィックス

`date: (日付指定)`
- 日付で検索します。以下の書式に対応しています。

1. 部分一致 (例: `date:202401`)
   2024年1月のノートを検索します。

2. 以降 (例: `date:>=20240101`)
   2024年1月1日以降のノートを検索します。

3. 以前 (例: `date:<=20240131`)
   2024年1月31日以前のノートを検索します。

4. 期間 (例: `date:20240101-20240131`)
   2024年1月1日から1月31日までのノートを検索します。

`time: (時刻指定)`
- 時刻 (`hhmmss` 形式) で検索します。
- `_` (アンダーバー) を「任意の一文字」ワイルドカードとして使えます。
- 6桁未満の入力は、右側がワイルドカードで埋められます。

  (例: `time:09` → `09____` → 9時台 (`09mmss`) にヒット)
  (例: `time:0900` → `0900__` → 9時00分台 (`0900ss`) にヒット)
  (例: `time:__30` → `__30__` → 毎時30分台 (`hh30mm`) にヒット)
  (例: `time:____00` → `____00` → 毎分00秒 (`hhmm00`) にヒット)

`is:orphan` (孤立ノート)
- どのノートからもリンクされておらず、どのノートへもリンクしていない「孤立したノート」を検索します。
- リンクのメンテナンスや、整理漏れの発見に役立ちます。
  (例: `is:orphan tag:アイデア` → 孤立しているアイデアノートを抽出)

----------------------------------------------------------------------------------------------------
■ グローバル検索 (プレフィックスなし)
(例: `Python`)

「本文・メモ検索」チェックボックスが...
- **OFF (デフォルト)**: `title:`, `tag:`, `key:`, `ikey:`, `date:`, `time:` \
を対象に検索します。
- **ON (低速)**: 上記に加え、`memo:` と `fulltext:` も対象に含めて検索します。
"""

        textbox = ctk.CTkTextbox(self, wrap="word")
        textbox.pack(fill="both", expand=True, padx=10, pady=(10, 5))
        textbox.insert("1.0", help_text)
        textbox.configure(state="disabled")

        close_button = ctk.CTkButton(self, text="閉じる", command=self.destroy)
        close_button.pack(pady=(0, 10), padx=10)

    def iconbitmap(self, *args, **kwargs):
        """
        iconbitmap の呼び出しをインターセプト（横取り）する。
        CustomTkinterが内部でアイコンをデフォルトに戻そうとしても、
        強制的にカスタムアイコンを設定し直す。
        """
        if self._custom_icon_path:
            try:
                # 常にカスタムアイコンパスを使って親メソッドを呼ぶ
                super().iconbitmap(self._custom_icon_path)
            except Exception:
                # ウィンドウが存在しない場合などのエラーを無視
                pass
        else:
            # カスタムアイコンがない場合は、通常の動作をさせる
            try:
                super().iconbitmap(*args, **kwargs)
            except Exception:
                pass


if __name__ == "__main__":
    app = Synapsen_Nexus()
    if app.icon_path:
        try:
            # 'default=' を指定し、OSダイアログ(エクスプローラ等)にも適用
            app.iconbitmap(default=str(app.icon_path))
        except Exception as e:
            logger.error(f"Icon default setting error: {e}")
    else:
        logger.warning(
            "警告: アイコンファイル (assets/synapsen.ico) が見つかりません。"
        )
    app.mainloop()

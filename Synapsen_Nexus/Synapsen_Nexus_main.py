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

# 分割したモジュールをインポート
import logging

from utils import (
    load_app_config,
    load_sql_data_file,
    open_pdf_viewer,
    build_memo_display,
    build_references_display,
    update_note_in_db,
    _update_note_links,
    delete_note_from_db,
    get_pdf_page_image,
)
from search_parser import parse_or_expression

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
    setup_logging("Synapsen_Normalisierer")
    logger = logging.getLogger("Normalisierer")  # このファイル用のロガー取得
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


class Synapsen_Nexus(ctk.CTk):
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

        self.grid_columnconfigure(0, weight=3)
        self.grid_columnconfigure(1, weight=2)
        self.grid_rowconfigure(1, weight=1)

        # --- アプリケーションの状態変数 ---
        self.df = None                      # ノートデータを保持するDataFrame
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
        self.filter_panel_expanded = False  # フィルターパネルが開いているか

        self.sort_ascending = (
            True  # ソート順を保持する変数 (デフォルトは (昇順/古い順))
        )

        self.selected_keys = set()  # 選択されたノートのKeyを保持するセット

        self._current_search_id = 0  # 検索リクエストのID
        self._search_lock = threading.Lock()

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

        # ウィンドウが初めて表示されたら on_map を呼ぶ
        self.bind("<Map>", self.on_map)
        # 最大化失敗時のフォールバックサイズ指定
        self.geometry("1200x800")  # (on_mapが呼ばれる前の初期サイズ)

        # ウィンドウを閉じるときのイベントをフック
        self.protocol("WM_DELETE_WINDOW", self.on_closing)

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
        """アプリ終了時の処理 (バックアップ作成 - 強制上書き)"""
        if self.loaded_db_path and Path(self.loaded_db_path).exists():
            try:
                # バックアップフォルダ作成
                db_path = Path(self.loaded_db_path)
                backup_dir = db_path.parent / "db_backups"
                backup_dir.mkdir(exist_ok=True)

                # 日付付きファイル名でコピー (例: Synapsen_Master_20231027.db)
                today = datetime.datetime.now().strftime("%Y%m%d")
                backup_path = backup_dir / f"{db_path.stem}_{today}.db"

                # 常に上書きコピー
                shutil.copy2(db_path, backup_path)
                print(f"DBバックアップを更新しました: {backup_path.name}")

            except Exception as e:
                print(f"DBバックアップ失敗: {e}")

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
                self._apply_default_search_query()
                self.perform_search()

        except FileNotFoundError as e:
            messagebox.showerror("設定エラー", str(e))
            self.destroy()
        except Exception as e:
            messagebox.showerror(
                "設定読み込みエラー", f"config.iniの読み込みに失敗しました: {e}"
            )
            self.destroy()

    def refresh_unique_tags(self):
        """
        タグリストを更新する。
        config設定(include_all_tags_for_autocomplete)によって挙動が変わる。
        """
        # 1. 事前定義タグでセットを初期化
        tags_set = set(self.predefined_tags)

        # ★変更: 設定が True の場合のみ、全ノートのタグをスキャンして追加
        if self.include_all_tags_for_autocomplete:
            if self.df is not None and not self.df.empty and "tags" in self.df.columns:
                valid_tags_series = self.df["tags"].dropna()
                valid_tags_series = valid_tags_series[valid_tags_series != ""]

                for tags_str in valid_tags_series:
                    current_note_tags = [
                        t.strip() for t in tags_str.split(";") if t.strip()
                    ]
                    tags_set.update(current_note_tags)

        # 3. ソートしてリスト化
        self.all_unique_tags = sorted(list(tags_set))
        logger.debug(
            "タグリスト更新 "
            f"(全タグ含む: {self.include_all_tags_for_autocomplete}): "
            f"{len(self.all_unique_tags)} 件"
        )

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

        # --- 右パネル (詳細・プレビュー) ---
        self.details_frame = ctk.CTkFrame(self)
        self.details_frame.grid(row=1, column=1, padx=(0, 10), pady=10, sticky="nsew")

        # グリッドの行設定 (プレビュー領域、メモ、引用元のために変更)
        self.details_frame.grid_rowconfigure(5, weight=1)  # PDFプレビュー (row 5)
        self.details_frame.grid_rowconfigure(7, weight=2)  # メモスクロール領域 (row 7)
        self.details_frame.grid_rowconfigure(
            9, weight=1
        )  # 引用元スクロール領域 (row 9)
        self.details_frame.grid_columnconfigure(1, weight=1)

        # 1. タイトル (row=0)
        ctk.CTkLabel(self.details_frame, text="タイトル:", anchor="w").grid(
            row=0, column=0, padx=10, pady=5, sticky="w"
        )
        self.title_label = ctk.CTkLabel(
            self.details_frame, text="", wraplength=300, justify="left", anchor="w"
        )
        self.title_label.grid(row=0, column=1, padx=10, pady=5, sticky="ew")

        # 2. キー (row=1)
        ctk.CTkLabel(self.details_frame, text="キー:", anchor="w").grid(
            row=1, column=0, padx=10, pady=5, sticky="w"
        )
        self.key_label = ctk.CTkLabel(self.details_frame, text="", anchor="w")
        self.key_label.grid(row=1, column=1, padx=10, pady=5, sticky="ew")

        # 3. インデックス キー (row=2)
        ctk.CTkLabel(self.details_frame, text="インデックス キー:", anchor="w").grid(
            row=2, column=0, padx=10, pady=5, sticky="w"
        )
        self.cpkey_label = ctk.CTkLabel(self.details_frame, text="", anchor="w")
        self.cpkey_label.grid(row=2, column=1, padx=10, pady=5, sticky="ew")

        # 4. タグ (row=3)
        ctk.CTkLabel(self.details_frame, text="タグ:", anchor="w").grid(
            row=3, column=0, padx=10, pady=5, sticky="w"
        )
        self.tags_label = ctk.CTkLabel(
            self.details_frame, text="", wraplength=300, justify="left", anchor="w"
        )
        self.tags_label.grid(row=3, column=1, padx=10, pady=5, sticky="ew")

        # 5. 概要 (Summary)
        ctk.CTkLabel(self.details_frame, text="概要:", anchor="w").grid(
            row=4, column=0, padx=10, pady=5, sticky="w"
        )
        self.summary_label = ctk.CTkLabel(
            self.details_frame, text="", wraplength=400, justify="left", anchor="w"
        )
        self.summary_label.grid(row=4, column=1, padx=10, pady=5, sticky="ew")

        # 6. PDFプレビュー (row=5)
        self.pdf_preview_label = ctk.CTkLabel(
            self.details_frame,
            text="ノートを選択するとプレビューが表示されます",
            fg_color="gray20",  # プレースホルダーの背景色
            anchor="center",
            text_color="gray70",  # プレースホルダーの文字色
        )
        self.pdf_preview_label.grid(
            row=5, column=0, columnspan=2, padx=10, pady=10, sticky="nsew"
        )

        # 7. メモ (ラベル) (row=6)
        ctk.CTkLabel(self.details_frame, text="メモ:", anchor="w").grid(
            row=6, column=0, padx=10, pady=5, sticky="nw"
        )

        # 8. メモ (スクロールフレーム) (row=7)
        self.memo_display_frame = ctk.CTkScrollableFrame(self.details_frame)
        self.memo_display_frame.grid(row=7, column=1, padx=10, pady=4, sticky="nsew")

        # 9. 引用元 (ラベル) (row=8)
        ctk.CTkLabel(self.details_frame, text="引用元:", anchor="w").grid(
            row=8, column=0, padx=10, pady=4, sticky="nw"
        )

        # 10. 引用元 (スクロールフレーム) (row=9)
        self.references_display_frame = ctk.CTkScrollableFrame(
            self.details_frame, label_text="このノートを引用しているノート"
        )
        self.references_display_frame.grid(
            row=9, column=1, padx=10, pady=5, sticky="nsew"
        )

        # 11. 編集ボタンのフレーム (row=10)
        self.edit_button_frame = ctk.CTkFrame(
            self.details_frame, fg_color="transparent"
        )
        self.edit_button_frame.grid(row=10, column=0, columnspan=2, pady=10)

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

        # ナビゲーションキーのリスト
        navigation_keys = (
            "Return",
            "Escape",
            "Left",
            "Up",
            "Down",
            "Right",
            "Home",
            "End",
        )

        if event.keysym in navigation_keys:
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

        last_tag_word = ""

        if match_value:
            # 'tag:abc' の 'abc' の部分 (group 2) を取得
            last_tag_word = match_value.group(2).strip()

        suggestions = []
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
            if i == self.selected_suggestion_index:
                fg_color = "gray30"
            else:
                fg_color = "transparent"

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

        num_suggestions = len(self.current_suggestions)
        if event.keysym == "Down":
            self.selected_suggestion_index = (
                self.selected_suggestion_index + 1
            ) % num_suggestions
        elif event.keysym == "Up":
            self.selected_suggestion_index = (
                self.selected_suggestion_index - 1 + num_suggestions
            ) % num_suggestions

        # 選択項目がリストに表示されるようにスクロール
        self.autocomplete_frame._parent_canvas.yview_moveto(
            self.selected_suggestion_index / num_suggestions
        )

        # 選択ハイライトを更新 (保存しておいた引数を使う)
        query, cursor_pos, match_value = self._last_suggestion_args
        self.show_autocomplete(self.current_suggestions, query, cursor_pos, match_value)
        return "break"  # 他のキーバインドを抑制

    def confirm_suggestion(self, event):
        """Enterキーで選択中の候補を確定する。"""
        if (
            self.autocomplete_frame.winfo_ismapped()
            and self.selected_suggestion_index != -1
            and self._last_suggestion_args is not None
        ):

            # 保存しておいた引数を取得
            query, cursor_pos, match_value = self._last_suggestion_args
            # select_suggestion を呼び出す
            self.select_suggestion(
                self.current_suggestions[self.selected_suggestion_index],
                query,
                cursor_pos,
                match_value,
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
                    icon_label = ctk.CTkLabel(
                        self.collapsed_icons_frame,
                        text=icon,
                        text_color=color,
                        font=("", 16),
                    )
                    icon_label.pack(side="left", padx=2)

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
                self.db_conn = None

            self.df = load_sql_data_file(filepath)  # utils (full_textなし)
            self.loaded_db_path = filepath

            # 読み取り専用のDB接続を保持 (メインのUI用)
            self.db_conn = sqlite3.connect(f"file:{filepath}?mode=ro", uri=True)

            # --- 2. テーブル構造の確認 (書き込み接続) ---
            # アプリ起動時/DB読込時に note_links テーブルが存在するか確認し、なければ作成
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

            # --- 3. UIリセット ---
            self.refresh_unique_tags()
            self.perform_search()

            # 変更したノートを再表示するロジック
            if key_to_redisplay and self.df is not None:
                # key_to_redisplay を使って、更新後のdfから行を検索
                target_note_row = self.df[self.df["key"] == key_to_redisplay]

                if not target_note_row.empty:
                    # 編集したノートが見つかったら、詳細を再表示
                    self.show_details(target_note_row.iloc[0])
                else:
                    # ノートが見つからない場合 (削除されたなど) はクリア
                    self.clear_details()
            else:
                # 再表示するキーが指定されていない場合はクリア
                self.clear_details()

            self.filter_panel_expanded = False
            self.sync_filter_panel_view()

        except sqlite3.OperationalError as e:
            messagebox.showerror(
                "データベース読み込みエラー",
                f"DB接続に失敗しました (読み取り専用モード): {e}",
            )
        except Exception as e:
            messagebox.showerror("データベース読み込みエラー", str(e))

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

            icon_label = ctk.CTkLabel(
                row_frame, text=icon, text_color=color, font=("", 16), width=20
            )
            icon_label.pack(side="left")

            cb = ctk.CTkCheckBox(
                row_frame,
                text=key,
                variable=var,
                onvalue="1",
                offvalue="0",
                command=self._trigger_search_now,  # チェック時に検索を再実行
            )
            cb.pack(side="left", expand=True, fill="x")

            self.filter_checkboxes[key] = var

    def perform_search(self):
        """
        【修正版】
        検索処理のエントリーポイント。
        UIスレッドで直接検索せず、バックグラウンドスレッドを開始します。
        """
        if self.df is None:
            self.update_results_list(pd.DataFrame())
            return

        # 1. UIからの検索条件取得（これはメインスレッドで行う必要がある）
        user_query = self.search_entry.get().strip()
        include_full_text = self.fts_checkbox.get()
        current_ascending = self.sort_ascending  # 現在のソート設定を取得

        # 除外タグのクエリ結合
        final_query = user_query

        # チェックボックスがON かつ 設定されたタグがある場合
        if self.exclude_tags_checkbox.get() == 1 and self.exclude_tags_by_default:
            # マイナス検索クエリを作成 (例: "-tag:Archive -tag:Done")
            exclusion_parts = [f"-tag:{tag}" for tag in self.exclude_tags_by_default]
            exclusion_query = " ".join(exclusion_parts)

            # ユーザーの入力がある場合は AND で結合、なければそのまま使用
            if final_query:
                final_query = f"({final_query}) AND ({exclusion_query})"
            else:
                final_query = exclusion_query

            logger.debug(
                f"除外タグ適用後の内部クエリ: {final_query}", extra={"sensitive": True}
            )

        # IndexKey フィルターの状態取得
        selected_filter_keys = []
        for key, var in self.filter_checkboxes.items():
            if var.get() == "1":
                selected_filter_keys.append(key)

        # 2. 検索IDを更新
        with self._search_lock:
            self._current_search_id += 1
            search_id = self._current_search_id

        # 3. UIの更新: 検索中表示
        self.results_list.configure(
            label_text="検索結果 (検索中...)", label_text_color="#E0a800"  # オレンジ色
        )

        # マウスカーソルを待機状態に
        self.configure(cursor="watch")

        # 4. バックグラウンドスレッドの開始
        thread = threading.Thread(
            target=self._execute_search_worker,
            args=(
                search_id,
                final_query,
                include_full_text,
                selected_filter_keys,
                self.loaded_db_path,
                current_ascending,
            ),
            daemon=True,
        )
        thread.start()

    def _execute_search_worker(
        self,
        search_id,
        query_text,
        include_full_text,
        selected_filter_keys,
        db_path,
        ascending_flag,
    ):
        """
        バックグラウンドスレッドで実行される検索ロジック。
        Args:
            search_id (int): この検索の一意なID。
            query_text (str): 検索クエリ文字列。
            include_full_text (bool): 本文・メモ検索を含めるか。
            selected_filter_keys (list): 選択されたIndexKeyフィルターのリスト。
            db_path (Path): FTS用のDBファイルパス。
            ascending_flag (bool): ソート順。DefaultはTrueで昇順。
        """
        try:
            # (A) ベースDF (メモリ上のメタデータ)
            base_df = self.df

            # (B) --- 1. IndexKey フィルター (Pandas) ---
            key_filter_mask = pd.Series([True] * len(base_df), index=base_df.index)
            if selected_filter_keys:
                key_filter_mask = base_df["commonplace_key"].isin(selected_filter_keys)

            # (C) FTS用のDB接続(conn)準備
            conn = None

            # 孤立ノート (is:orphan) 検索の判定
            is_orphan_search = "is:orphan" in query_text.lower()
            if is_orphan_search:
                # is:orphan をクエリから除去して、他の検索語(もしあれば)と併用できるようにする
                # (例: "is:orphan tag:Python" -> Pythonタグを持つ孤立ノート)
                query_text = re.sub(
                    r"is:orphan", "", query_text, flags=re.IGNORECASE
                ).strip()

            # '本文・メモ検索' (include_full_text) が有効 又は
            # 'memo:'/'fulltext:'/'text:' プレフィックスがクエリに含まれる場合に
            # DB接続(conn)を準備する
            query_lower = query_text.lower()
            needs_db_search = (
                include_full_text
                or "memo:" in query_lower
                or "fulltext:" in query_lower
                or "text:" in query_lower
                or is_orphan_search
            )

            if needs_db_search:  # クエリが空でもorphan検索なら接続する
                try:
                    # スレッドごとに読み取り専用接続を作成
                    conn = sqlite3.connect(f"file:{db_path}?mode=ro", uri=True)
                except Exception as e:
                    logger.error(f"FTS用DB接続エラー: {e}")
                    # conn は None のまま

            # (D) --- 2. 検索クエリ (search_parser.py) ---
            query_mask = pd.Series([True] * len(base_df), index=base_df.index)
            if query_text:
                try:
                    query_mask = parse_or_expression(
                        base_df, query_text, include_full_text, conn
                    )
                except Exception as e:
                    logger.error(f"検索クエリ解析エラー: {e}")
                    query_mask = pd.Series([False] * len(base_df), index=base_df.index)

            # (E) --- 孤立ノートフィルタリング ---
            orphan_mask = pd.Series([True] * len(base_df), index=base_df.index)

            if is_orphan_search and conn:
                try:
                    # リンクテーブルにある全ての source_key と target_key を取得
                    cursor = conn.cursor()
                    cursor.execute(
                        """
                        SELECT source_key FROM note_links
                        UNION
                        SELECT target_key FROM note_links
                    """
                    )
                    linked_keys = {row[0] for row in cursor.fetchall()}

                    # リンクされているキーに含まれ「ない」ものを True にする
                    orphan_mask = ~base_df["key"].isin(linked_keys)

                except Exception as e:
                    logger.error(f"孤立ノート検索エラー: {e}")
                    # エラー時は全除外等の安全策をとるか、無視して続行するか
                    orphan_mask = pd.Series([False] * len(base_df), index=base_df.index)

            # --- 3. クリーンアップ ---
            if conn:
                conn.close()

            # --- 4. 【重要】最終的な絞り込み ---
            final_mask = key_filter_mask & query_mask & orphan_mask
            filtered_df = base_df[final_mask]

            if not filtered_df.empty:
                # 受け取った ascending_flag を使用してソート
                filtered_df = filtered_df.sort_values(
                    by="key", ascending=ascending_flag
                )

            # --- 5. メインスレッドに結果を渡す ---
            self.after(0, lambda: self._on_search_complete(search_id, filtered_df))

        except Exception as e:
            logger.error(f"検索スレッドエラー: {e}", exc_info=True)
            self.after(0, lambda: self._on_search_complete(search_id, pd.DataFrame()))

    def _on_search_complete(self, search_id, result_df):
        """
        【新規】
        検索完了時にメインスレッドから呼び出されるコールバック。
        結果の検証とUI更新を行います。
        """
        # マウスカーソルを元に戻す
        self.configure(cursor="")

        # 最新の検索IDと一致するか確認
        with self._search_lock:
            if search_id != self._current_search_id:
                return

        # キャッシュの更新
        self.filtered_df_cache = result_df

        # UIリストの更新
        self.update_results_list(result_df)
        self.update_collapsed_filter_view()

        self.results_list.configure(label_text_color=("gray10", "#DCE4EE"))

        # デバッグ用出力
        logger.debug(f"検索完了 (ID: {search_id}): {len(result_df)} 件ヒット")

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
        現在の検索結果（または全データ）からランダムに1件のノートを選択し、
        詳細ペインに表示する。
        """

        # 1. 使用するDataFrameを選択
        target_df = None
        if self.filtered_df_cache is not None and not self.filtered_df_cache.empty:
            # 現在の検索結果キャッシュが存在すれば、それを使用
            target_df = self.filtered_df_cache
        elif self.df is not None and not self.df.empty:
            # 検索結果が空（または未検索）だが、DB自体は読み込まれている場合
            target_df = self.df

        # 2. 選択対象が存在するかチェック
        if target_df is None or target_df.empty:
            messagebox.showinfo(
                "ランダムノート",
                "表示できるノートがありません。\nデータベースを読み込んでいるか確認してください。",
                parent=self,
            )
            return

        try:
            # 3. DataFrameからランダムに1行を取得
            # .sample() はランダムにDataFrameを返し、
            # .iloc[0] でその最初の行 (Series) を取得
            random_note_row = target_df.sample(n=1).iloc[0]

            # 4. 詳細表示メソッドを呼び出す
            self.show_details(random_note_row)

        except Exception as e:
            logger.error(f"ランダムノートの表示中にエラー: {e}")
            messagebox.showerror(
                "エラー", f"ノートのランダム表示に失敗しました:\n{e}", parent=self
            )

    def open_canvas(self):
        """キャンバスウィンドウを開く"""
        if hasattr(self, "canvas_window") and self.canvas_window.winfo_exists():
            self.canvas_window.focus()
            return
        self.canvas_window = CanvasWindow(self)

    # --- UI更新・表示メソッド ---
    def update_results_list(self, df_to_show):
        """
        フィルタリングされたDataFrameに基づき、検索結果リストUIを更新する。

        Args:
            df_to_show (pd.DataFrame): リストに表示するデータ。
        """
        for widget in self.results_list.winfo_children():
            widget.destroy()

        self.results_list.configure(label_text=f"検索結果 ({len(df_to_show)}件)")

        for index, row in df_to_show.iterrows():
            item_frame = ctk.CTkFrame(self.results_list, fg_color="transparent")
            item_frame.pack(fill="x", padx=5, pady=2)

            # チェックボックス ---
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

            # --- イベントバインド ---
            def create_show_details_handler(note_row=row):
                def handler(event):
                    self.show_details(note_row)

                return handler

            def create_open_pdf_handler(note_row=row):
                def handler(event):
                    self.open_pdf(note_row)

                return handler

            # シングルクリックで詳細表示
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

    def toggle_note_selection(self, key, var):
        """チェックボックスの切り替え時の処理"""
        if var.get() == "on":
            self.selected_keys.add(key)
        else:
            self.selected_keys.discard(key)

        self.update_selection_ui_state()

    def clear_selection(self):
        """選択をすべて解除する"""
        self.selected_keys.clear()
        self.update_selection_ui_state()
        # リストのチェックボックス表示を更新するため、現在の検索結果でリストを再描画
        if self.filtered_df_cache is not None:
            self.update_results_list(self.filtered_df_cache)

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
        if not self.selected_keys or self.df is None:
            return

        # 選択されたキーに対応する行を取得
        selected_df = self.df[self.df["key"].isin(self.selected_keys)]

        if selected_df.empty:
            return

        # 日付順（またはKey順）にソートしてリスト化
        selected_df = selected_df.sort_values(by="key")

        link_texts = []
        for _, row in selected_df.iterrows():
            key = row["key"]
            title = row["title"]
            # Synapsenのリンク形式: [[Key: Title]]
            link_texts.append(f"[[{key}: {title}]]")

        # 改行区切りで結合
        clipboard_text = "\n".join(link_texts)

        # クリップボードへコピー
        self.clipboard_clear()
        self.clipboard_append(clipboard_text)
        self.update()  # クリップボード更新を確定させるために必要

        messagebox.showinfo(
            "コピー完了",
            f"{len(link_texts)}件のリンクをクリップボードにコピーしました。\n"
            "新しいノートのメモ欄にペーストして、MOC（目次）として利用できます。",
            parent=self,
        )

    def show_selected_graph(self):
        """
        選択されたノート(self.selected_keys)のみでグラフを表示する。
        """
        if not self.selected_keys:
            return

        if self.df is None:
            return

        # 全データ(self.df)から、選択されたキーを持つ行だけを抽出
        selected_df = self.df[self.df["key"].isin(self.selected_keys)]

        if selected_df.empty:
            messagebox.showinfo("情報", "選択されたノートのデータが見つかりません。")
            return

        logger.info(f"Selected Graph: {len(selected_df)} notes")
        # 既存のグラフ生成メソッドを再利用
        self.generate_and_show_graph(target_df=selected_df)

    def clear_details(self):
        """詳細表示ペインの内容をすべてクリアする。"""
        self.title_label.configure(text="")
        self.key_label.configure(text="")
        self.cpkey_label.configure(text="")
        self.tags_label.configure(text="")

        # --- PDFプレビューのクリア ▼ ---
        self.preview_image_object = None

        # 既存のラベルを破棄
        if hasattr(self, "pdf_preview_label"):
            self.pdf_preview_label.destroy()

        # ラベルを「クリア状態」で再作成
        self.pdf_preview_label = ctk.CTkLabel(
            self.details_frame,
            text="ノートを選択するとプレビューが表示されます",
            fg_color="gray20",  # プレースホルダーの背景色
            anchor="center",
            text_color="gray70",  # プレースホルダーの文字色
        )
        self.pdf_preview_label.grid(
            row=4, column=0, columnspan=2, padx=10, pady=10, sticky="nsew"
        )
        # -----------------------------------

        # 選択中ノートとボタンの状態をクリア
        self.current_selected_row = None
        self.edit_button.configure(state="disabled")
        self.delete_button.configure(state="disabled")
        self.open_preview_button.configure(state="disabled")

        # memo_display_frame内のすべてのウィジェット（ラベル）を削除
        for widget in self.memo_display_frame.winfo_children():
            widget.destroy()

        # references_display_frame内もクリア
        for widget in self.references_display_frame.winfo_children():
            widget.destroy()
        self.references_display_frame.configure(
            label_text="このノートを引用しているノート"
        )

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
                if memo_data and memo_data[0]:
                    current_memo = memo_data[0]
                else:
                    current_memo = ""

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
        ★ 新規追加:
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
        ui_master: キャンバス等、別のウィンドウから呼び出す場合にそのウィンドウを指定
        """
        if self.df is None or self.db_conn is None:
            messagebox.showwarning("データなし", "データベースが読み込まれていません。")
            return

        # 1. メタデータを self.df から取得 (高速)
        target_note_row = self.df[self.df["key"] == key]
        if target_note_row.empty:
            messagebox.showwarning(
                "ノート不明", f"ID '{key}' に一致するノートが見つかりませんでした。"
            )
            return

        # 2. memo と full_text を DB から取得
        try:
            cursor = self.db_conn.cursor()
            cursor.execute(
                "SELECT memo, full_text, summary FROM notes WHERE key = ?", (key,)
            )
            db_data = cursor.fetchone()
            note_data = target_note_row.iloc[0].copy()
            if db_data:
                note_data["memo"] = db_data[0]
                note_data["full_text"] = db_data[1]
                note_data["summary"] = db_data[2]
        except Exception as e:
            logger.error(f"プレビュー用のDBデータ取得エラー: {e}")
            note_data = target_note_row.iloc[0]

        # 3. プレビューウィンドウのインスタンスを作成 (ui_masterを渡す)
        preview_win = NotePreviewWindow(
            self, note_data, default_view_mode, ui_master=ui_master
        )
        preview_win.focus()

    def show_details(self, row_data):
        """
        選択されたノートの詳細を右ペインに表示する。

        Args:
            row_data (pd.Series): 表示するノートの行データ。
        """
        if not isinstance(row_data, pd.Series):
            logger.error(
                f"Error: show_details に不正なデータ型が渡されました: {type(row_data)}"
            )
            self.clear_details()
            return

        # 選択中の行データを保持し、ボタンを有効化
        self.current_selected_row = row_data
        self.edit_button.configure(state="normal")
        self.delete_button.configure(state="normal")

        row = row_data
        self.summary_label.configure(text=row.get("summary", ""))
        self.title_label.configure(text=row.get("title", ""))
        self.key_label.configure(text=row.get("key", ""))
        self.cpkey_label.configure(text=row.get("commonplace_key", ""))
        self.open_preview_button.configure(state="normal")

        # タグ表示（文字列をリストに変換して表示）
        tags_str = str(row.get("tags", ""))
        tags_list = [tag for tag in tags_str.split(";") if tag]
        self.tags_label.configure(text=", ".join(tags_list))

        # --- ▼ PDFプレビューの表示 ▼ ---

        # 既存のプレビューラベルを破棄
        if hasattr(self, "pdf_preview_label"):
            self.pdf_preview_label.destroy()
        max_preview_width = 225
        pil_image = get_pdf_page_image(
            row_data,
            self.loaded_db_path,
            self.pdf_root_folder,
            max_width=max_preview_width,
            pdf_archive_folder=self.pdf_archive_folder,
        )

        if pil_image:
            # Pillow Image を CTkImage に変換
            self.preview_image_object = ctk.CTkImage(
                light_image=pil_image,
                dark_image=pil_image,
                size=(pil_image.width, pil_image.height),
            )
            # プレビュー用ラベルを「画像付き」で再作成
            self.pdf_preview_label = ctk.CTkLabel(
                self.details_frame,
                image=self.preview_image_object,
                text="",  # テキストを削除
                fg_color="transparent",  # 背景を透明に
            )
        else:
            # 画像取得失敗
            self.preview_image_object = None
            # プレビュー用ラベルを「失敗状態」で再作成
            self.pdf_preview_label = ctk.CTkLabel(
                self.details_frame,
                image=None,
                text="プレビューの読み込みに失敗しました",
                fg_color="gray20",
                text_color="#D9534F",
            )

        # 新しく作成したラベルをグリッドに配置
        self.pdf_preview_label.grid(
            row=5, column=0, columnspan=2, padx=10, pady=10, sticky="nsew"
        )

        # --- メモと引用元の取得 ---

        current_key = row.get("key", "")

        memo_text = ""
        backlinks_df = pd.DataFrame()  # 空で初期化

        try:
            # (db_conn は読み取り専用接続)
            cursor = self.db_conn.cursor()

            # 1. このノートの 'memo' と 'summary' を取得
            cursor.execute(
                "SELECT memo, summary FROM notes WHERE key = ?", (current_key,)
            )
            data = cursor.fetchone()
            if data:
                memo_text = str(data[0])
                # メモリ上にある row_data にも最新の memo と summary をセット
                self.current_selected_row["memo"] = memo_text
                self.current_selected_row["summary"] = str(data[1])
                self.summary_label.configure(text=str(data[1]))

            # 2. 引用元 (Backlinks) を 'note_links' テーブルから取得

            # リンクテーブルを検索
            sql = "SELECT source_key FROM note_links WHERE target_key = ?"

            cursor.execute(sql, (current_key,))

            matching_keys = {row[0] for row in cursor.fetchall()}

            if matching_keys:
                logger.debug(
                    f"[show_details] リンクテーブル検索ヒット: {len(matching_keys)} 件 "
                    f"(Key: {current_key})"
                )
                backlinks_df = self.df[self.df["key"].isin(matching_keys)]
            else:
                logger.debug(
                    f"[show_details] リンクテーブル検索 0件 " f"(Key: {current_key})"
                )

        except Exception as e:
            logger.error(f"詳細表示のためのDBアクセスエラー: {e}", exc_info=True)
            pass  # 失敗時は空のまま

        # メモ表示（リンク構築）
        for widget in self.memo_display_frame.winfo_children():
            widget.destroy()
        frame_width = 450
        build_memo_display(
            self.memo_display_frame,
            memo_text,
            self.df,
            lambda key: self.open_preview_window(key, default_view_mode="compact"),
            frame_width,
        )
        # 引用元UIを構築
        build_references_display(
            self.references_display_frame,
            backlinks_df,
            lambda key: self.open_preview_window(key, default_view_mode="compact"),
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

        Args:
            key (str): PDFを開くノートの 'key' (ID)。
        """
        if self.df is None:
            messagebox.showerror("エラー", "CSVデータが読み込まれていません。")
            return

        target_note_row = self.df[self.df["key"] == key]

        if target_note_row.empty:
            messagebox.showwarning(
                "ノート不明", f"ID '{key}' に一致するノートが見つかりませんでした。"
            )
            return

        note_data = target_note_row.iloc[0]
        self.open_pdf(note_data)

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
        # 1. データフレーム決定
        df = target_df if target_df is not None else self.filtered_df_cache

        if df is None or df.empty:
            if not output_path:
                messagebox.showinfo(
                    "グラフ表示", "グラフ化するノートがありません。", parent=self
                )
            return

        # 2. パフォーマンス制限
        if len(df) > 500 and not output_path:
            messagebox.showwarning(
                "グラフ表示", "件数が多いため500件に制限します。", parent=self
            )
            df = df.head(500)

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
            logger.error(f"Graph error: {e}")
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
        現在詳細ペインで選択されているノートと、
        そのノートに「リンクしている」または「リンクされている」ノートのみでグラフを表示する。
        """
        if not center_key:
            messagebox.showinfo("情報", "グラフの中心となるノートのKeyがありません。")
            return
        if self.df is None or self.db_conn is None:
            return

        related_keys = set()
        related_keys.add(center_key)

        try:
            # (db_conn は読み取り専用接続)
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

        except Exception as e:
            logger.error(f"ローカルグラフのリンク取得エラー: {e}")
            # エラーが発生しても、中心ノードのみでグラフ生成を試みる

        # 2. 関連キーのみを含むDataFrameを作成
        local_df = self.df[self.df["key"].isin(related_keys)]

        if local_df.empty:
            messagebox.showinfo("情報", "関連するノートが見つかりませんでした。")
            return

        # 3. グラフ生成 (target_df を指定して呼び出し)
        logger.info(f"Local Graph: {len(local_df)} notes related to {center_key}")
        self.generate_and_show_graph(target_df=local_df)

    # --- エクスポート機能 ---
    def export_search_results(self, include_pdf=False):
        """
        検索結果(または選択中)のデータをエクスポートする。
        include_pdf=True の場合、同じフォルダに統合PDFも生成する。
        """
        target_df = None
        mode = "search_results"

        # 対象データの決定
        if self.selected_keys:
            if self.df is not None:
                target_df = self.df[self.df["key"].isin(self.selected_keys)]
                mode = "selected_items"
        else:
            target_df = self.filtered_df_cache

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
            messagebox.showerror(
                "エクスポートエラー", f"失敗しました:\n{e}", parent=self
            )

    def merge_and_export_pdf(self):
        """メニューから「統合PDF」単体を選んだ場合のラッパー"""
        # 対象データの決定
        target_df = None
        if self.selected_keys:
            if self.df is not None:
                target_df = self.df[self.df["key"].isin(self.selected_keys)]
        else:
            target_df = self.filtered_df_cache

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

        # ExportManager の merge_pdf を直接利用
        if self.exporter.merge_pdf(
            target_df,
            Path(save_path),
            self.pdf_root_folder,
            pdf_archive_folder=self.pdf_archive_folder,
        ):
            messagebox.showinfo("完了", f"PDFを保存しました:\n{save_path}", parent=self)
        else:
            messagebox.showwarning(
                "失敗", "結合可能なPDFが見つかりませんでした。", parent=self
            )

    def create_moc_markdown(self):
        """MOC (Markdown) 生成処理のラッパー"""
        # 対象データの決定 (選択中 > 検索結果)
        target_df = None
        if self.selected_keys:
            if self.df is not None:
                target_df = self.df[self.df["key"].isin(self.selected_keys)]
        else:
            target_df = self.filtered_df_cache

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
        if self.exporter.generate_moc_markdown(
            target_df, Path(save_path), self.loaded_db_path
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

        if self.df is None or self.db_conn is None:
            messagebox.showwarning("データなし", "データベースが読み込まれていません。")
            return

        key_to_edit = note_data.get("key")

        # --- DBから最新の 'memo' と 'summary' を取得 ---
        # NoteEditorWindowに渡す note_data (pd.Series) をコピーして更新する
        note_data_with_memo = note_data.copy()
        try:
            cursor = self.db_conn.cursor()
            cursor.execute(
                "SELECT memo, summary FROM notes WHERE key = ?", (key_to_edit,)
            )
            db_data = cursor.fetchone()
            if memo_data := db_data:
                note_data_with_memo["memo"] = str(memo_data[0])
                note_data_with_memo["summary"] = str(memo_data[1])
            else:
                note_data_with_memo["memo"] = ""
                note_data_with_memo["summary"] = ""
        except Exception as e:
            logger.error(f"編集ウィンドウ用のメモ取得エラー: {e}")
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
            f"以下のノートをマスターデータベースから完全に削除しますか？\n\n"
            f"Key: {key_to_delete}\n"
            f"Title: {title_to_delete}\n\n"
            f"この操作は元に戻せません。",
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

        self.title("検索ヘルプ")
        self.geometry("550x620")
        self.transient(parent_app)
        self.grab_set()

        help_text = """
Synapsen Nexus 検索クエリ リファレンス

■ 基本
- 検索語をスペースで区切ると `AND` 検索になります。
  (例: `Type_Permanent 薬物動態学`)
- `OR` を使用すると `OR` 検索ができます。
  (例: `Type_Fleeting OR Question`)
- `()` でグループ化できます。
  (例: `(tag:Type_Fleeting OR tag:Question) AND (ikey:学習 OR ikey:タスク)`)
- `-` (ハイフン) を検索語の前に付けると `NOT` 検索になります。
  (例: `衛生 -memo:古い`)

---
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

---
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

---
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

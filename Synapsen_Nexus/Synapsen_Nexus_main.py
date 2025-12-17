import shutil
import customtkinter as ctk
import threading
from tkinter import filedialog, messagebox
import pandas as pd
from pathlib import Path
import sys
import datetime
import sqlite3
import subprocess
import platform

# --- ロガー設定 ---
current_dir = Path(__file__).parent
root_dir = current_dir.parent
if str(root_dir) not in sys.path:
    sys.path.append(str(root_dir))

try:
    from logging_setup import setup_logging

    setup_logging("Synapsen_Nexus")
    import logging

    logger = logging.getLogger("Nexus")
except ImportError:
    print("Warning: logging_setup.py not found.")
    import logging

    logger = logging.getLogger()

# --- 依存モジュールのインポート ---
# mixinsフォルダを確実にパスに通す
mixins_path = current_dir / "mixins"
if str(current_dir) not in sys.path:
    sys.path.append(str(current_dir))

try:
    # 実行環境によってインポート元が変わる可能性があるため柔軟に対応
    try:
        from Synapsen_Nexus.mixins import (
            NexusDatabaseMixin,
            NexusUiMixin,
            NexusSearchMixin,
        )
        from Synapsen_Nexus.utils import (
            load_app_config,
            open_pdf_viewer,
            build_memo_display,
            build_references_display,
            update_note_in_db,
            _update_note_links,
            delete_note_from_db,
            get_pdf_page_image,
            find_file_in_paths,
        )
        from Synapsen_Nexus.list_navigator import ListNavigatorMixin
        from Synapsen_Nexus.saved_search_manager import SavedSearchManager
        from Synapsen_Nexus.graph_manager import GraphManager
        from Synapsen_Nexus.export_manager import ExportManager
        from Synapsen_Nexus.canvas_window import CanvasWindow
        from Synapsen_Nexus.preview_window import NotePreviewWindow
        from Synapsen_Nexus.editor_window import NoteEditorWindow
        from theme import SemanticColors as Colors

    except ImportError:
        # ディレクトリ直下で実行された場合など
        from mixins import NexusDatabaseMixin, NexusUiMixin, NexusSearchMixin
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
        )
        from list_navigator import ListNavigatorMixin
        from saved_search_manager import SavedSearchManager
        from graph_manager import GraphManager
        from export_manager import ExportManager
        from canvas_window import CanvasWindow
        from preview_window import NotePreviewWindow
        from editor_window import NoteEditorWindow
        from theme import SemanticColors as Colors

except ImportError as e:
    logger.critical(f"Critical Import Error: {e}")
    messagebox.showerror("起動エラー", f"モジュールの読み込みに失敗しました:\n{e}")
    sys.exit(1)


class Synapsen_Nexus(
    ctk.CTk, NexusDatabaseMixin, NexusUiMixin, NexusSearchMixin, ListNavigatorMixin
):
    """
    Synapsen Nexus メインアプリケーション
    Mixinを使用して機能を分割・整理しています。
    """

    def __init__(self):
        """アプリケーションを初期化し、ウィンドウと変数をセットアップする。"""
        super().__init__()
        self.icon_path = self.get_icon_path()
        self.title("Synapsen Nexus")
        self.configure(
            fg_color=(Colors.BACKGROUND_HOLLOW, Colors.BACKGROUND_DARK_HOLLOW)
        )

        self.grid_columnconfigure(0, weight=3)  # 左パネル
        self.grid_columnconfigure(1, weight=2)  # 右パネル
        self.grid_rowconfigure(1, weight=1)

        # --- アプリケーションの状態変数 ---
        # 変数の初期化 (エラー回避のため先に定義)
        self.pdf_root_folder = None  # 統合PDFが存在する(メイン)フォルダのルートパス
        self.pdf_archive_folder = (
            None  # 統合PDFが存在する(アーカイブフォルダ等)サブフォルダのルートパス
        )
        self.nexus_output_folder = Path("Nexus_Output")
        self.browser_path = None

        # 空のデフォルト値をセット
        self.key_icons = {}
        self.key_colors = {}
        self.commonplace_keys_options = []
        self.predefined_tags = []
        self.include_all_tags_for_autocomplete = True
        self.exclude_tags_by_default = []

        self.config_data = {}  # Canvas等から参照される設定辞書
        self.filter_checkboxes = {}  # IndexKeyフィルターのチェックボックス変数
        self.filter_panel_expanded = False  # 左フィルターパネルが開いているか
        self.details_panel_expanded = True  # 右詳細パネルが開いているか (初期表示)

        self.sort_ascending = (
            True  # ソート順を保持する変数 (デフォルトは (昇順/古い順))
        )
        self.selected_keys = set()  # 選択されたノートのKeyを保持するセット
        self.current_selected_row = None

        self.search_timer = None  # デバウンス（検索遅延）用タイマー
        self._details_timer = None  # リストのクリックのタイマー
        self._preview_request_id = 0  # プレビュー生成リクエスト管理ID

        # CTkImageオブジェクトへの参照を保持 (ガベージコレクション対策)
        self.preview_image_object = None

        # --- ページネーション管理 ---
        self.current_page = 0
        self.items_per_page = 50
        self.total_items = 0
        self.current_where_clause = ""
        self.current_params = []
        self._pending_reveal_key = None  # ページジャンプ後のカーソル移動予約

        # --- Mixinの初期化 ---
        self.setup_database_variables()  # DatabaseMixin
        self.setup_search_variables()  # SearchMixin
        self.setup_navigation_variables()  # ListNavigatorMixin

        # --- マネージャクラスの初期化 ---
        self.search_manager = SavedSearchManager(self)

        # --- 設定読み込み ---
        self.load_config()

        # --- UI構築 ---
        self.create_widgets()  # UiMixin

        # 検索マネージャの読み込み
        root_path = (
            self.base_path if getattr(sys, "frozen", False) else self.base_path.parent
        )
        self.search_manager.load_saved_searches(root_path)

        # ExportManager (UI構築後、config情報を渡す)
        self.exporter = ExportManager(
            {"key_icons": self.key_icons, "key_colors": self.key_colors}
        )

        # UI反映 (フィルタアイコンなど)
        self.populate_key_filters()
        self.sync_filter_panel_view()
        self.update_details_panel(None)

        # --- DB接続 ---
        default_db_path = self.config_data.get("database_path")
        if default_db_path and default_db_path.is_file():
            self.load_db_from_path(default_db_path)
        else:
            self.perform_search()  # DBなしでもUI初期化

        # ショートカット & イベント
        self._setup_shortcuts()
        self.bind("<Map>", self.on_map)
        self.protocol("WM_DELETE_WINDOW", self.on_closing)

        # フォールバックサイズ
        self.geometry("1200x800")

    def on_map(self, event):
        """
        ウィンドウが初めて画面に描画されたときに呼び出される。
        ここで最大化を実行する。
        """
        try:
            self.unbind("<Map>")
            self.state("zoomed")
        except Exception:
            pass

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
        except Exception:
            pass
        return None

    def load_config(self):
        """config.ini読み込みと変数へのセット"""
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

            # 設定読み込み試行
            self.config_data = load_app_config(self.base_path)

            # 読み込み成功時のみ上書き
            self.pdf_root_folder = self.config_data.get("pdf_root_folder", Path(""))
            self.pdf_archive_folder = self.config_data.get("pdf_archive_folder", None)
            self.nexus_output_folder = self.config_data.get(
                "nexus_output_folder", Path("Nexus_Output")
            )
            self.browser_path = self.config_data.get("browser_path", None)
            self.key_icons = self.config_data.get("key_icons", {})
            self.key_colors = self.config_data.get("key_colors", {})
            self.commonplace_keys_options = self.config_data.get(
                "commonplace_keys_options", []
            )
            self.predefined_tags = self.config_data.get("predefined_tags", [])

            self.include_all_tags_for_autocomplete = self.config_data.get(
                "include_all_tags_for_autocomplete", True
            )

            self.exclude_tags_by_default = self.config_data.get(
                "exclude_tags_by_default", []
            )

        except FileNotFoundError:
            # configがない場合は警告を出すが、デフォルト値で続行する
            messagebox.showwarning(
                "設定ファイルなし",
                "config.ini が見つかりませんでした。\nデフォルト設定で起動します。",
            )
        except Exception as e:
            messagebox.showerror("設定エラー", f"設定の読み込みに失敗しました:\n{e}")

    def on_closing(self):
        """終了処理: バックアップとクリーンアップ"""
        if self.loaded_db_path and Path(self.loaded_db_path).exists():
            try:
                db_path = Path(self.loaded_db_path)
                backup_dir = db_path.parent / "db_backups"
                backup_dir.mkdir(exist_ok=True)
                today_str = datetime.datetime.now().strftime("%Y%m%d")
                backup_filename = f"{db_path.stem}_{today_str}.db"
                shutil.copy2(db_path, backup_dir / backup_filename)

                # ローテーション
                backup_files = sorted(list(backup_dir.glob(f"{db_path.stem}_*.db")))
                if len(backup_files) > 7:
                    for old in backup_files[:-7]:
                        if old.name != backup_filename:
                            old.unlink()
            except Exception as e:
                logger.error(f"Backup error: {e}")

        if self.db_conn:
            self.db_conn.close()
        self.destroy()

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

        # ページ切り替え (左右キー)
        self.bind("<Left>", lambda e: self._handle_shortcut(self.prev_page))
        self.bind("<Right>", lambda e: self._handle_shortcut(self.next_page))

        # 最初/最後のページへ移動 (Alt + Home/End)
        self.bind("<Alt-Home>", lambda e: self._handle_shortcut(self.first_page))
        self.bind("<Alt-End>", lambda e: self._handle_shortcut(self.last_page))

        # Escキー: フォーカス解除
        self.bind("<Escape>", self._handle_escape)

        # Ctrl+F: 検索バーへフォーカス (Find)
        self.bind("<Control-f>", self._focus_search)
        self.bind("<Control-F>", self._focus_search)

        # --- 3. リスト操作用ショートカット ---
        self.setup_navigation_shortcuts()  # ListNavigatorMixin

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
        focused = self.focus_get()
        # 入力ウィジェットにフォーカスがある場合は実行しない
        if focused and focused.winfo_class() in ["Entry", "Text"]:
            return

        # 入力中でなければコマンドを実行
        command()

    def _focus_search(self, event):
        """検索バーにフォーカスを移動し、全選択状態にする"""
        self.search_entry.focus_set()
        self.search_entry.select_range(0, "end")
        return "break"  # デフォルトの動作を抑制

    # -------------------------------------------------------------------------

    def populate_key_filters(self):
        """IndexKeyフィルタのチェックボックス生成"""
        for widget in self.key_filter_frame.winfo_children():
            widget.destroy()
        self.filter_checkboxes.clear()

        for key in self.commonplace_keys_options:
            var = ctk.StringVar(value="0")
            row = ctk.CTkFrame(self.key_filter_frame, fg_color="transparent")
            row.pack(anchor="w", padx=10, pady=2, fill="x")

            icon = self.key_icons.get(key.lower(), "•")

            raw_color = self.key_colors.get(key.lower(), "gray")
            if raw_color != "gray" and not raw_color.startswith("#"):
                color = f"#{raw_color}"
            else:
                color = raw_color

            ctk.CTkLabel(
                row, text=icon, text_color=color, font=("", 16), width=20
            ).pack(side="left")
            ctk.CTkCheckBox(
                row,
                text=key,
                variable=var,
                onvalue="1",
                offvalue="0",
                command=self._trigger_search_now,
                fg_color=Colors.UI_BASIC,
                hover_color=Colors.adjust_brightness(Colors.UI_BASIC),
                checkmark_color=Colors.adjust_brightness(Colors.UI_BASIC, 1.8),
            ).pack(side="left", expand=True, fill="x")
            self.filter_checkboxes[key] = var

    def sync_filter_panel_view(self):
        if self.filter_panel_expanded:
            self.key_filter_frame.grid()
            self.toggle_filter_button.configure(text="▼ IndexKey")
        else:
            self.key_filter_frame.grid_remove()
            self.toggle_filter_button.configure(text="▶ IndexKey")
        self.update_collapsed_filter_view()

    def update_collapsed_filter_view(self):
        for w in self.collapsed_icons_frame.winfo_children():
            w.destroy()
        if not self.filter_panel_expanded:
            selected = [k for k, v in self.filter_checkboxes.items() if v.get() == "1"]
            if not selected:
                ctk.CTkLabel(self.collapsed_icons_frame, text="", font=("", 16)).pack(
                    side="left"
                )
            else:
                for k in selected:
                    icon = self.key_icons.get(k.lower(), "•")

                    raw_col = self.key_colors.get(k.lower(), "gray")
                    col = (
                        f"#{raw_col}"
                        if (raw_col != "gray" and not raw_col.startswith("#"))
                        else raw_col
                    )

                    ctk.CTkLabel(
                        self.collapsed_icons_frame,
                        text=icon,
                        text_color=col,
                        font=("", 16),
                    ).pack(side="left", padx=2)

    def toggle_filter_panel(self):
        self.filter_panel_expanded = not self.filter_panel_expanded
        self.sync_filter_panel_view()

    def toggle_details_panel(self):
        if self.details_panel_expanded:
            self.details_frame.grid_remove()
            self.details_panel_expanded = False
            self.toggle_details_button.configure(text="◀ 詳細")
        else:
            self.details_frame.grid()
            self.details_panel_expanded = True
            self.toggle_details_button.configure(text="▶ 詳細")

    # --- アクション系メソッド ---
    def open_canvas(self, background=False, notes_to_add=None):
        if hasattr(self, "canvas_window") and self.canvas_window.winfo_exists():
            if not background:
                if self.canvas_window.state() == "iconic":
                    self.canvas_window.deiconify()
                self.canvas_window.lift()
                self.canvas_window.focus_force()
            if notes_to_add:
                self.canvas_window.add_selected_notes(target_keys=notes_to_add)
            return
        self.canvas_window = CanvasWindow(
            self, background=background, initial_notes=notes_to_add
        )

    def refresh_unique_tags(self):
        """タグリスト更新"""
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

    def _reload_db(self):
        if self.loaded_db_path:
            self.load_db_from_path(self.loaded_db_path)

    def load_database_dialog(self):
        """「目次データベースを開く」ボタンの動作。ファイルダイアログを開く。"""
        filepath = filedialog.askopenfilename(
            title="目次データベースファイルを選択",
            filetypes=[("SQLite Database", "*.db"), ("All files", "*.*")],
        )
        if not filepath:
            return
        self.load_db_from_path(Path(filepath))

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

    def send_selection_to_canvas(self):
        """
        現在選択中のノートをCanvasに追加する。
        Canvasが開いていない場合は自動で開く。
        フォーカスはNexusに残す（連続作業用）。
        """
        if not self.selected_keys:
            # 選択がない場合、ラベルを一瞬赤くして通知
            self.selection_info_label.configure(
                text="選択なし!", text_color=Colors.LABEL_DANGER
            )
            self.after(1000, self.update_selection_ui_state)  # 1秒後に元に戻す
            return

        # 1. Canvasをバックグラウンドモードで開く (またはノートを渡す)
        self.open_canvas(background=True, notes_to_add=self.selected_keys)

        # 2. 成功メッセージ
        count = len(self.selected_keys)
        self.selection_info_label.configure(
            text=f"Canvasに追加: {count}件", text_color=Colors.LABEL_SUCCESS
        )
        self.after(1500, self.update_selection_ui_state)

        # 3. フォーカスをNexusに維持
        self.focus_force()

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
            self.selection_info_label.configure(
                text_color=Colors.adjust_brightness(Colors.LABEL_WARNING)
            )
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
        メイン画面で選択中の全ノートに対し、リンクを追記する。
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
            from datetime import datetime

            now_str = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

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
                    "UPDATE notes SET memo = ?, updated_at = ? WHERE key = ?",
                    (new_memo, now_str, key),
                )

                # 6. リンクテーブルも更新
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
                    text_color=(Colors.TEXT_LINK, Colors.TEXT_LINK_BRIGHT),
                    hover_color=Colors.adjust_brightness(Colors.BACKGROUND_PANEL, 0.9),
                    anchor="w",
                    height=20,
                    command=lambda k=cp_key: self.quick_search(f"ikey:{k}"),
                ).pack(side="left", fill="x", expand=True)
            else:
                ctk.CTkLabel(self.text_info_frame, text="-", anchor="w").grid(
                    row=r, column=1, sticky="ew"
                )

        add_row("Index Key:", create_ikey_widget)

        # 4. ファイル
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
                    fg_color=Colors.UI_PREVIEW,
                    hover_color=Colors.adjust_brightness(Colors.UI_PREVIEW),
                    text_color="black",
                    command=lambda p=target_open_path: self.open_file_location(p),
                ).pack(side="left", padx=(0, 5))
            if display_filename != "(不明)":
                ctk.CTkButton(
                    f,
                    text=display_filename,
                    fg_color="transparent",
                    text_color=(Colors.TEXT_LINK, Colors.TEXT_LINK_BRIGHT),
                    hover_color=Colors.adjust_brightness(Colors.BACKGROUND_PANEL, 0.9),
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
                        fg_color=Colors.adjust_brightness(Colors.BACKGROUND_PANEL, 0.9),
                        text_color=(
                            Colors.TEXT_LINK,
                            Colors.TEXT_LINK_BRIGHT,
                        ),
                        hover_color=Colors.adjust_brightness(Colors.BACKGROUND_PANEL),
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

    # --- エクスポート機能 ---
    def export_search_results(self, include_pdf=False):
        """
        検索結果(または選択中)のデータをエクスポートする。
        """
        # ヘルパーを使ってDataFrameを取得
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

    # --- リスト関連 ---
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
                        text="リスト外",
                        text_color=Colors.adjust_brightness(Colors.LABEL_WARNING),
                    )
                    self.after(1500, self.update_selection_ui_state)
                else:
                    # ページ遷移を実行
                    self.selection_info_label.configure(
                        text="ジャンプ中...",
                        text_color=Colors.adjust_brightness(Colors.LABEL_WARNING),
                    )

                    self._pending_reveal_key = target_key  # 読み込み完了後の予約
                    self.current_page = target_page
                    self.perform_search(reset_page=False)  # 指定ページで再検索
            else:
                # 検索条件に合致しない (除外タグなど)
                self.selection_info_label.configure(
                    text="リスト外",
                    text_color=Colors.adjust_brightness(Colors.LABEL_WARNING),
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

    # --- デバウンス関連 ---
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

    def quick_search(self, query):
        """指定されたクエリを検索バーに入力して即座に検索を実行する"""
        self.search_entry.delete(0, "end")
        self.search_entry.insert(0, query)
        self.perform_search()


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

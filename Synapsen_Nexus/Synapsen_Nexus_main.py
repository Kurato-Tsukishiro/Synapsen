import customtkinter as ctk
from tkinter import filedialog, messagebox
import pandas as pd
from pathlib import Path
import re
import sys
import webbrowser
import networkx as nx
from pyvis.network import Network
import datetime
import shutil
from textwrap import dedent
from pypdf import PdfReader, PdfWriter

# 分割したモジュールをインポート
from utils import (
    load_app_config, load_sql_data_file, open_pdf_viewer,
    build_memo_display, build_references_display, find_backlinks_df,
    update_note_in_db, delete_note_from_db,
    get_pdf_page_image,
    get_pdf_uri_for_note
)
from search_parser import parse_or_expression

from preview_window import NotePreviewWindow
from editor_window import NoteEditorWindow
from saved_search_manager import SavedSearchManager


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
        self.pdf_root_folder = None         # config.iniから読み込むPDFのルートパス
        self.loaded_db_path = None          # 現在開いているDBのパス

        self.key_icons = {}                 # IndexKeyごとのアイコン
        self.key_colors = {}                # IndexKeyごとの色
        self.commonplace_keys_options = []  # IndexKeyの全オプション

        self.filter_checkboxes = {}         # IndexKeyフィルターのチェックボックス変数
        self.filter_panel_expanded = False  # フィルターパネルが開いているか

        self.selected_keys = set()          # 選択されたノートのKeyを保持するセット

        self.filtered_df_cache = pd.DataFrame()
        self.current_selected_row = None

        # CTkImageオブジェクトへの参照を保持 (ガベージコレクション対策)
        self.preview_image_object = None

        # --- オートコンプリート関連 ---
        self.predefined_tags = []          # オートコンプリート用のタグリスト
        self.selected_suggestion_index = -1
        self.current_suggestions = []
        self.search_timer = None           # デバウンス（検索遅延）用タイマー
        self.suggestion_timer = None       # オートコンプリート用タイマー
        self._last_suggestion_args = None  # 予測変換の引数(query, cursor_pos, match)

        self.base_path = None  # アプリの基準パス (config.ini と同じ場所)

        # 検索マネージャのインスタンス化
        self.search_manager = SavedSearchManager(self)

        self.create_widgets()
        self.load_config()

        # ウィンドウが初めて表示されたら on_map を呼ぶ
        self.bind("<Map>", self.on_map)
        # 最大化失敗時のフォールバックサイズ指定
        self.geometry("1200x800")  # (on_mapが呼ばれる前の初期サイズ)

    def on_map(self, event):
        """
        ウィンドウが初めて画面に描画されたときに呼び出される。
        ここで最大化を実行する。
        """
        try:
            self.unbind("<Map>")
            self.state('zoomed')
            print("[DEBUG] ウィンドウを最大化しました。")
        except Exception as e:
            print(f"ウィンドウの最大化に失敗しました: {e}")

    def get_icon_path(self):
        """
        実行環境(.exe or .py)に応じて、
        プロジェクトルートの 'assets' フォルダにある
        'synapsen.ico' のパスを返す。
        """
        try:
            if getattr(sys, 'frozen', False):
                # .exe実行の場合 (exeと同じフォルダがプロジェクトルート)
                project_root = Path(sys.executable).parent
            else:
                # .pyスクリプト実行の場合 (このファイルの親フォルダがプロジェクトルート)
                project_root = Path(__file__).parent.parent

            icon_path = project_root / 'assets' / 'synapsen.ico'

            if icon_path.is_file():
                return icon_path
        except Exception as e:
            print(f"Error finding icon path: {e}")
        return None

    def load_config(self):
        """
        config.iniファイルからアプリケーション設定を読み込み、適用する。
        """
        try:
            # 実行ファイルのパスを基準にconfig.iniを探す
            if getattr(sys, 'frozen', False):
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
            self.pdf_root_folder = config_data.get('pdf_root_folder', Path(''))
            self.key_icons = config_data.get('key_icons', {})
            self.key_colors = config_data.get('key_colors', {})
            self.commonplace_keys_options = config_data.get(
                'commonplace_keys_options', []
                )
            self.predefined_tags = config_data.get('predefined_tags', [])

            # フィルターチェックボックスをUIに反映
            self.populate_key_filters()

            # 検索マネージャの読み込み
            try:
                # search_manager は config.ini と同じ場所(root) のパスを必要とする
                if getattr(sys, 'frozen', False):
                    # .exe の場合、self.base_path (e.g., dist/) が root
                    root_path = self.base_path
                else:
                    # .py の場合、self.base_path.parent (e.g., Synapsen/) が root
                    root_path = self.base_path.parent

                self.search_manager.load_saved_searches(root_path)
            except Exception as e:
                print(f"保存済み検索の読み込みエラー: {e}")

            # デフォルトDBが設定されていれば自動で読み込む
            default_db_path = config_data.get('database_path')
            if default_db_path and default_db_path.is_file():
                self.load_db_from_path(default_db_path)
            else:
                if default_db_path:
                    print(f"デフォルトデータベースが見つかりません: {default_db_path}")
                self.perform_search()  # 空の状態で検索を実行

        except FileNotFoundError as e:
            messagebox.showerror("設定エラー", str(e))
            self.destroy()
        except Exception as e:
            messagebox.showerror("設定読み込みエラー", f"config.iniの読み込みに失敗しました: {e}")
            self.destroy()

    def create_widgets(self):
        # --- トップフレーム ---
        top_frame = ctk.CTkFrame(self)
        top_frame.grid(
            row=0, column=0, columnspan=2, padx=10, pady=(10, 0), sticky="ew"
            )
        top_frame.grid_columnconfigure(1, weight=1)

        # ボタンを左側にまとめるフレーム
        left_button_frame = ctk.CTkFrame(top_frame, fg_color="transparent")
        left_button_frame.pack(side="left", padx=5)

        ctk.CTkButton(
            left_button_frame, text="DBを開く", command=self.load_database_dialog
        ).pack(side="left", padx=(0, 5))

        # 検索バーコンテナ
        search_container = ctk.CTkFrame(top_frame, fg_color="transparent")
        search_container.pack(side="left", fill="x", expand=True, padx=5)

        self.search_entry = ctk.CTkEntry(
            search_container,
            placeholder_text=(
                "検索 (例: (date:>=20240101 AND date:<=20240131) AND" +
                " (tag:Python OR tag:C#))"
            )
        )
        self.search_entry.pack(fill="x")

        # 右側ボタンフレーム
        right_button_frame = ctk.CTkFrame(top_frame, fg_color="transparent")
        right_button_frame.pack(side="right", padx=(5, 10))

        # 本文検索(FTS)
        self.fts_checkbox = ctk.CTkCheckBox(
            right_button_frame, text="本文検索"
        )
        self.fts_checkbox.pack(side="left", padx=5)
        self.fts_checkbox.configure(command=self._trigger_search_now)

        # 選択数表示ラベル
        self.selection_info_label = ctk.CTkLabel(
            right_button_frame,
            text="選択: 0",
            font=("", 12, "bold"),
            text_color="gray",
            width=60                # 固定幅を確保してレイアウト揺れを防ぐ
        )
        self.selection_info_label.pack(side="left", padx=(10, 0))

        # グラフメニュー
        self.graph_menu_var = ctk.StringVar(value="グラフ表示")

        self.graph_menu = ctk.CTkOptionMenu(
            right_button_frame,
            variable=self.graph_menu_var,
            values=["全体 (Global)", "関連 (Local)", "選択 (Selected)"],
            command=self.handle_graph_menu,
            width=140,
            fg_color="#585a9c",
            button_color="#494B83"
        )
        self.graph_menu.pack(side="left", padx=(5, 0))

        # リンクコピーボタン
        self.copy_links_button = ctk.CTkButton(
            right_button_frame,
            text="リンクコピー",
            command=self.copy_selected_links,
            width=90,
            fg_color="#28a745",    # 緑色 (コピー系のアクション色)
            hover_color="#218838",
            state="disabled"
        )
        self.copy_links_button.pack(side="left", padx=(5, 0))

        # エクスポートメニュー
        self.export_menu_var = ctk.StringVar(value="エクスポート")
        self.export_menu = ctk.CTkOptionMenu(
            right_button_frame,
            variable=self.export_menu_var,
            values=["データ (CSV/TXT)", "統合PDF (Merge)", "全て (Data + PDF)"],
            command=self.handle_export_menu,
            width=140,
            fg_color="#17a2b8",  # シアン系 (出力・情報アクションとして区別)
            button_color="#138496"
        )
        self.export_menu.pack(side="left", padx=(5, 0))

        # 選択解除ボタン
        self.clear_selection_button = ctk.CTkButton(
            right_button_frame,
            text="×",
            command=self.clear_selection,
            width=30,
            fg_color="#6C757D",
            hover_color="#5A6268"
        )
        self.clear_selection_button.pack(side="left", padx=(2, 0))

        # ランダムノートボタン
        self.random_note_button = ctk.CTkButton(
            right_button_frame,
            text="閃き (R)",
            command=self.show_random_note,
            width=80,
            fg_color="#585a9c",
            hover_color="#494B83"
        )
        self.random_note_button.pack(side="left", padx=(5, 0))

        # スマート検索UI
        smart_search_frame = ctk.CTkFrame(top_frame, fg_color="transparent")
        smart_search_frame.pack(side="right", padx=5)

        self.save_search_button = ctk.CTkButton(
            smart_search_frame,
            text="検索保存",
            command=self.search_manager.save_current_search,
            width=80
        )
        self.save_search_button.pack(side="left", padx=(5, 0))

        # 保存済み検索呼び出しボタン
        self.saved_search_combo = ctk.CTkComboBox(
            smart_search_frame,
            values=["保存済み検索..."],
            width=150,
            command=self.search_manager.on_saved_search_selected
        )
        self.saved_search_combo.pack(side="left", padx=5)
        self.saved_search_combo.set("保存済み検索...")

        # 検索バーのイベントバインド
        self.search_entry.bind("<KeyRelease>", self.handle_keyrelease)
        self.search_entry.bind("<FocusOut>", self.hide_autocomplete)
        self.search_entry.bind("<FocusIn>", self.schedule_suggestions)
        self.search_entry.bind("<Down>", self.navigate_suggestions)
        self.search_entry.bind("<Up>", self.navigate_suggestions)
        self.search_entry.bind("<Return>", self.confirm_suggestion)

        # オートコンプリート用の非表示フレーム
        self.autocomplete_frame = ctk.CTkScrollableFrame(self, label_text="")

        # --- 左パネル ---
        self.left_panel = ctk.CTkFrame(self, fg_color="transparent")
        self.left_panel.grid(row=1, column=0, padx=10, pady=10, sticky="nsew")
        self.left_panel.grid_rowconfigure(2, weight=1)
        self.left_panel.grid_columnconfigure(0, weight=1)

        # フィルターコンテナ
        filter_container = ctk.CTkFrame(self.left_panel)
        filter_container.grid(row=0, column=0, sticky="ew")
        filter_container.grid_columnconfigure(1, weight=1)
        self.toggle_filter_button = ctk.CTkButton(
            filter_container, text="", command=self.toggle_filter_panel,
            width=20
        )
        self.toggle_filter_button.grid(row=0, column=0, padx=5, pady=5)

        # フィルター非表示時に選択中アイコンを表示するフレーム
        self.collapsed_icons_frame = ctk.CTkFrame(
            filter_container, fg_color="transparent"
            )
        self.collapsed_icons_frame.grid(
            row=0, column=1, padx=5, pady=5, sticky="w"
            )

        # IndexKeyフィルターのスクロールフレーム (初期非表示)
        self.key_filter_frame = ctk.CTkScrollableFrame(
            self.left_panel, label_text=""
            )
        self.key_filter_frame.grid(
            row=1, column=0, padx=0, pady=(0, 5), sticky="nsew"
            )

        # 検索結果リスト
        self.results_list = ctk.CTkScrollableFrame(
            self.left_panel, label_text="ノート一覧"
            )
        self.results_list.grid(row=2, column=0, padx=0, pady=0, sticky="nsew")

        # --- 右パネル ---
        self.details_frame = ctk.CTkFrame(self)
        self.details_frame.grid(
            row=1, column=1, padx=(0, 10), pady=10, sticky="nsew"
            )

        # グリッドの行設定 (プレビュー領域、メモ、引用元のために変更)
        self.details_frame.grid_rowconfigure(4, weight=1)  # PDFプレビュー
        self.details_frame.grid_rowconfigure(6, weight=2)  # メモ欄 (重み2)
        self.details_frame.grid_rowconfigure(8, weight=1)  # 引用元欄 (重み1)
        self.details_frame.grid_columnconfigure(1, weight=1)

        ctk.CTkLabel(
            self.details_frame, text="タイトル:", anchor="w"
            ).grid(row=0, column=0, padx=10, pady=5, sticky="w")
        self.title_label = ctk.CTkLabel(
            self.details_frame, text="",
            wraplength=300, justify="left", anchor="w"
            )
        self.title_label.grid(row=0, column=1, padx=10, pady=5, sticky="ew")

        ctk.CTkLabel(
            self.details_frame, text="キー:", anchor="w"
            ).grid(row=1, column=0, padx=10, pady=5, sticky="w")
        self.key_label = ctk.CTkLabel(self.details_frame, text="", anchor="w")
        self.key_label.grid(row=1, column=1, padx=10, pady=5, sticky="ew")

        ctk.CTkLabel(
            self.details_frame, text="インデックス キー:", anchor="w"
            ).grid(row=2, column=0, padx=10, pady=5, sticky="w")
        self.cpkey_label = ctk.CTkLabel(
            self.details_frame, text="", anchor="w"
            )
        self.cpkey_label.grid(row=2, column=1, padx=10, pady=5, sticky="ew")

        ctk.CTkLabel(
            self.details_frame, text="タグ:", anchor="w"
            ).grid(row=3, column=0, padx=10, pady=5, sticky="w")
        self.tags_label = ctk.CTkLabel(
            self.details_frame, text="",
            wraplength=300, justify="left", anchor="w"
            )
        self.tags_label.grid(row=3, column=1, padx=10, pady=5, sticky="ew")

        # 4. PDFプレビュー
        self.pdf_preview_label = ctk.CTkLabel(
            self.details_frame,
            text="ノートを選択するとプレビューが表示されます",
            fg_color="gray20",  # プレースホルダーの背景色
            anchor="center",
            text_color="gray70"  # プレースホルダーの文字色
        )
        self.pdf_preview_label.grid(
            row=4, column=0, columnspan=2, padx=10, pady=10, sticky="nsew"
        )

        # 5. メモ (ラベル)
        ctk.CTkLabel(
            self.details_frame, text="メモ:", anchor="w"
            ).grid(row=5, column=0, padx=10, pady=5, sticky="nw")

        # 6. メモ (スクロールフレーム)
        self.memo_display_frame = ctk.CTkScrollableFrame(self.details_frame)
        self.memo_display_frame.grid(
            row=6, column=1, padx=10, pady=5, sticky="nsew"
            )

        # 7. 引用元 (ラベル)
        ctk.CTkLabel(
            self.details_frame, text="引用元:", anchor="w"
            ).grid(row=7, column=0, padx=10, pady=5, sticky="nw")

        # 8. 引用元 (スクロールフレーム)
        self.references_display_frame = ctk.CTkScrollableFrame(
            self.details_frame, label_text="このノートを引用しているノート"
            )
        self.references_display_frame.grid(
            row=8, column=1, padx=10, pady=5, sticky="nsew"
            )

        # 9. 編集ボタン
        self.edit_button_frame = ctk.CTkFrame(
            self.details_frame, fg_color="transparent")
        self.edit_button_frame.grid(row=9, column=0, columnspan=2, pady=10)

        self.edit_button = ctk.CTkButton(
            self.edit_button_frame,
            text="このノートを編集",
            command=self.open_edit_dialog,
            state="disabled"
        )
        self.edit_button.pack(side="left", padx=10)

        self.delete_button = ctk.CTkButton(
            self.edit_button_frame,
            text="DBから削除",
            command=self.confirm_delete_note,
            fg_color="#D9534F",  # 赤色
            hover_color="#C9302C",  # 濃い赤色
            state="disabled"
        )
        self.delete_button.pack(side="left", padx=10)

        # フィルターパネルの初期表示を同期
        self.sync_filter_panel_view()

    # --- グラフメニューのハンドラ ---
    def handle_graph_menu(self, choice):
        """グラフメニューで選択された項目に応じて処理を分岐する"""
        if choice == "全体 (Global)":
            self.generate_and_show_graph()
        elif choice == "関連 (Local)":
            self.show_local_graph()
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

        self.export_menu.set("エクスポート")

    # --- オートコンプリート関連メソッド ---
    def handle_keyrelease(self, event):
        """検索バーでのキー入力（リリース）イベントを処理する。"""
        if event.keysym in ("Up", "Down", "Return", "Escape"):
            return
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

        # (match_value が None の場合 = 'tag:' と入力した直後)
        # last_tag_word は "" (空文字) となる

        suggestions = []
        if last_tag_word == "":
            # 'tag:' と入力した直後の場合は、全タグリストを表示
            suggestions = self.predefined_tags
        else:
            # 'tag:Py' のように入力中の場合は、前方一致検索
            last_word_lower = last_tag_word.lower()
            suggestions = [
                tag for tag in self.predefined_tags
                if tag.lower().startswith(last_word_lower)
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
                self.autocomplete_frame, text=suggestion, fg_color=fg_color,
                text_color=ctk.ThemeManager.theme["CTkLabel"]["text_color"],
                anchor="w",
                # select_suggestion に必要な情報をラムダで渡す
                command=lambda s=suggestion: self.select_suggestion(
                    s, query, cursor_pos, match_value
                )
            )
            btn.pack(fill="x", padx=5, pady=2)

        # 検索バーの真下に配置 (元のコード)
        x = self.search_entry.winfo_rootx() - self.winfo_rootx()
        y = (
            self.search_entry.winfo_rooty() -
            self.winfo_rooty() +
            self.search_entry.winfo_height()
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
            prefix_part = query[:match_value.start(2)]
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
        if (not self.autocomplete_frame.winfo_ismapped() or
                not self.current_suggestions or
                self._last_suggestion_args is None):
            return

        num_suggestions = len(self.current_suggestions)
        if event.keysym == "Down":
            self.selected_suggestion_index = (
                self.selected_suggestion_index + 1) % num_suggestions
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
        self.show_autocomplete(
            self.current_suggestions, query, cursor_pos, match_value
        )
        return "break"  # 他のキーバインドを抑制

    def confirm_suggestion(self, event):
        """Enterキーで選択中の候補を確定する。"""
        if (self.autocomplete_frame.winfo_ismapped() and
                self.selected_suggestion_index != -1 and
                self._last_suggestion_args is not None):

            # 保存しておいた引数を取得
            query, cursor_pos, match_value = self._last_suggestion_args
            # select_suggestion を呼び出す
            self.select_suggestion(
                self.current_suggestions[self.selected_suggestion_index],
                query, cursor_pos, match_value
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
                    if var.get() == '1':
                        selected_keys.append(key)

            if not selected_keys:
                ctk.CTkLabel(
                    self.collapsed_icons_frame, text="", font=("", 16)
                ).pack(side="left")
            else:
                for key in selected_keys:
                    icon = self.key_icons.get(key.lower(), '•')
                    color = self.key_colors.get(key.lower(), 'gray')
                    icon_label = ctk.CTkLabel(
                        self.collapsed_icons_frame, text=icon,
                        text_color=color, font=("", 16)
                        )
                    icon_label.pack(side="left", padx=2)

    # --- データ読み込み・検索実行メソッド ---
    def load_database_dialog(self):
        """「目次データベースを開く」ボタンの動作。ファイルダイアログを開く。"""
        filepath = filedialog.askopenfilename(
            title="目次データベースファイルを選択",
            filetypes=[("SQLite Database", "*.db"), ("All files", "*.*")]
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
            # utilsの関数でDataFrameを読み込む
            self.df = load_sql_data_file(filepath)
            self.loaded_db_path = filepath

            # UIをリセット・更新
            self.perform_search()

            # 変更したノートを再表示するロジック
            if key_to_redisplay and self.df is not None:
                # key_to_redisplay を使って、更新後のdfから行を検索
                target_note_row = self.df[self.df['key'] == key_to_redisplay]

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

        except Exception as e:
            messagebox.showerror("データベース読み込みエラー", str(e))

    def populate_key_filters(self):
        for widget in self.key_filter_frame.winfo_children():
            widget.destroy()
        self.filter_checkboxes.clear()

        for key in self.commonplace_keys_options:
            var = ctk.StringVar(value='0')
            row_frame = ctk.CTkFrame(
                self.key_filter_frame, fg_color="transparent"
                )
            row_frame.pack(anchor="w", padx=10, pady=2, fill="x")

            icon = self.key_icons.get(key.lower(), '•')
            color = self.key_colors.get(key.lower(), 'gray')

            icon_label = ctk.CTkLabel(
                row_frame, text=icon, text_color=color,
                font=("", 16), width=20
                )
            icon_label.pack(side="left")

            cb = ctk.CTkCheckBox(
                row_frame, text=key, variable=var,
                onvalue='1', offvalue='0',
                command=self._trigger_search_now  # チェック時に検索を再実行
            )
            cb.pack(side="left", expand=True, fill="x")

            self.filter_checkboxes[key] = var

    def perform_search(self):
        """
        現在のフィルター状態と検索クエリに基づき、DataFrameをフィルタリングし、
        結果リストを更新する。search_parser.parse_or_expression を使用する。
        """
        if self.df is None:
            self.update_results_list(pd.DataFrame())
            return

        filtered_df = self.df.copy()

        # 1. IndexKey フィルターを適用
        selected_keys = []
        for key, var in self.filter_checkboxes.items():
            if var.get() == '1':
                selected_keys.append(key)

        if selected_keys:
            filtered_df = (
                filtered_df[filtered_df['commonplace_key'].isin(selected_keys)]
            )

        # 2. 検索クエリを適用
        query_text = self.search_entry.get().strip()
        include_full_text = self.fts_checkbox.get()  # 本文検索が有効かを取得

        if query_text:
            try:
                # 2. 最初の関数 parse_or_expression にフラグを渡す
                final_mask = parse_or_expression(
                    filtered_df, query_text, include_full_text
                )
                filtered_df = filtered_df[final_mask]
            except Exception as e:
                print(f"検索クエリの解析エラー: {e}")
                # エラー時は空の結果を表示
                filtered_df = filtered_df.iloc[0:0]

        self.filtered_df_cache = filtered_df

        self.update_results_list(filtered_df)
        self.update_collapsed_filter_view()

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
        tag_value_pattern = r'(?i)(^|\s|\()tags?:([^\s\)]*)$'
        match_value = re.search(tag_value_pattern, query_to_cursor)

        # 正規表現: 'tag:' (または 'tags:') を入力した直後か
        tag_prefix_pattern = r'(?i)(^|\s|\()tags?:\s*$'
        match_prefix = re.search(tag_prefix_pattern, query_to_cursor)

        if match_prefix or match_value:
            # 'tag:' が入力されている場合のみ、予測変換を実行 (200ms後)
            # update_suggestions に必要な情報を渡す
            self.suggestion_timer = self.after(
                200,
                lambda q=query, c=cursor_pos, m=match_value:
                    self.update_suggestions(q, c, m)
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
        if (
                self.filtered_df_cache is not None and
                not self.filtered_df_cache.empty
        ):
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
                parent=self
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
            print(f"ランダムノートの表示中にエラー: {e}")
            messagebox.showerror(
                "エラー", f"ノートのランダム表示に失敗しました:\n{e}", parent=self)

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
            item_frame = ctk.CTkFrame(
                self.results_list, fg_color="transparent"
                )
            item_frame.pack(fill="x", padx=5, pady=2)

            # チェックボックス ---
            note_key = row.get('key')
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
                command=on_toggle
            )
            checkbox.pack(side="left", padx=(0, 5))

            cp_key = str(row.get("commonplace_key", "")).lower()
            icon = self.key_icons.get(cp_key, '•')
            color = self.key_colors.get(cp_key, 'gray')

            icon_label = ctk.CTkLabel(
                item_frame,
                text=icon,
                text_color=color,
                font=("", 16), width=20
                )
            icon_label.pack(side="left")

            display_text = f"[{row.get('date')}] {row.get('title', 'N/A')}"
            text_label = ctk.CTkLabel(
                item_frame, text=display_text, anchor="w")
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
        selected_df = self.df[self.df['key'].isin(self.selected_keys)]

        if selected_df.empty:
            return

        # 日付順（またはKey順）にソートしてリスト化
        selected_df = selected_df.sort_values(by='key')

        link_texts = []
        for _, row in selected_df.iterrows():
            key = row['key']
            title = row['title']
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
            parent=self
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
        selected_df = self.df[self.df['key'].isin(self.selected_keys)]

        if selected_df.empty:
            messagebox.showinfo("情報", "選択されたノートのデータが見つかりません。")
            return

        print(f"Selected Graph: {len(selected_df)} notes")
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
        if hasattr(self, 'pdf_preview_label'):
            self.pdf_preview_label.destroy()

        # ラベルを「クリア状態」で再作成
        self.pdf_preview_label = ctk.CTkLabel(
            self.details_frame,
            text="ノートを選択するとプレビューが表示されます",
            fg_color="gray20",   # プレースホルダーの背景色
            anchor="center",
            text_color="gray70"  # プレースホルダーの文字色
        )
        self.pdf_preview_label.grid(
            row=4, column=0, columnspan=2, padx=10, pady=10, sticky="nsew"
        )
        # -----------------------------------

        # 選択中ノートとボタンの状態をクリア
        self.current_selected_row = None
        self.edit_button.configure(state="disabled")
        self.delete_button.configure(state="disabled")

        # memo_display_frame内のすべてのウィジェット（ラベル）を削除
        for widget in self.memo_display_frame.winfo_children():
            widget.destroy()

        # references_display_frame内もクリア
        for widget in self.references_display_frame.winfo_children():
            widget.destroy()
        self.references_display_frame.configure(
            label_text="このノートを引用しているノート"
            )

    def open_preview_window(self, key):
        """
        指定されたキーのノートを新しい「簡易プレビュー」ウィンドウで開く。
        (メモ欄の [[key]] リンククリック時の動作)

        Args:
            key (str): 表示するノートの 'key' (ID)。
        """
        if self.df is None:
            messagebox.showwarning("データなし", "データベースが読み込まれていません。")
            return

        target_note_row = self.df[self.df['key'] == key]

        if target_note_row.empty:
            messagebox.showwarning("ノート不明", f"ID '{key}' に一致するノートが見つかりませんでした。")
            return

        note_data = target_note_row.iloc[0]

        # プレビューウィンドウ (読み取り専用) のインスタンスを作成
        preview_win = NotePreviewWindow(self, note_data)
        preview_win.focus()  # ウィンドウにフォーカスを当てる

    def show_details(self, row_data):
        """
        選択されたノートの詳細を右ペインに表示する。

        Args:
            row_data (pd.Series): 表示するノートの行データ。
        """
        if not isinstance(row_data, pd.Series):
            print(f"Error: show_details に不正なデータ型が渡されました: {type(row_data)}")
            self.clear_details()
            return

        # 選択中の行データを保持し、ボタンを有効化
        self.current_selected_row = row_data
        self.edit_button.configure(state="normal")
        self.delete_button.configure(state="normal")

        row = row_data
        self.title_label.configure(text=row.get('title', ''))
        self.key_label.configure(text=row.get('key', ''))
        self.cpkey_label.configure(text=row.get('commonplace_key', ''))

        # タグ表示（文字列をリストに変換して表示）
        tags_str = str(row.get('tags', ''))
        tags_list = [tag for tag in tags_str.split(';') if tag]
        self.tags_label.configure(text=", ".join(tags_list))

        # --- ▼ PDFプレビューの表示 ▼ ---

        # 既存のプレビューラベルを破棄
        if hasattr(self, 'pdf_preview_label'):
            self.pdf_preview_label.destroy()

        max_preview_width = 225  # プレビュー表示の最大幅

        pil_image = get_pdf_page_image(
            row_data,
            self.loaded_db_path,
            self.pdf_root_folder,
            max_width=max_preview_width
        )

        if pil_image:
            # Pillow Image を CTkImage に変換
            self.preview_image_object = ctk.CTkImage(
                light_image=pil_image,
                dark_image=pil_image,
                size=(pil_image.width, pil_image.height)
            )
            # プレビュー用ラベルを「画像付き」で再作成
            self.pdf_preview_label = ctk.CTkLabel(
                self.details_frame,
                image=self.preview_image_object,
                text="",                # テキストを削除
                fg_color="transparent"  # 背景を透明に
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
                text_color="#D9534F"  # 赤色
            )

        # 新しく作成したラベルをグリッドに配置
        self.pdf_preview_label.grid(
            row=4, column=0, columnspan=2, padx=10, pady=10, sticky="nsew"
        )

        # メモ表示（リンク構築）
        # (row=6, column=1 の memo_display_frame を使用)
        for widget in self.memo_display_frame.winfo_children():
            widget.destroy()

        memo_text = str(row.get('memo', ''))
        frame_width = 450  # 詳細ペインのメモ欄の幅

        build_memo_display(
            self.memo_display_frame,
            memo_text,
            self.df,
            self.open_preview_window,  # リンククリック時のコールバック
            frame_width
        )

        # 引用元の検索と表示
        # (row=8, column=1 の references_display_frame を使用)
        current_key = row.get('key', '')

        # 引用元DFを取得
        backlinks_df = find_backlinks_df(self.df, current_key)

        # 引用元UIを構築
        build_references_display(
            self.references_display_frame,
            backlinks_df,
            self.open_preview_window,
            self.key_icons,
            self.key_colors
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

        target_note_row = self.df[self.df['key'] == key]

        if target_note_row.empty:
            messagebox.showwarning("ノート不明", f"ID '{key}' に一致するノートが見つかりませんでした。")
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
            self.pdf_root_folder
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
        # 1. 使用するデータフレームを決定
        if target_df is not None:
            df = target_df
        else:
            # 引数がない場合は、「現在の(キャッシュされた)検索結果」を使用
            df = self.filtered_df_cache

        if df is None or df.empty:
            if not output_path:
                messagebox.showinfo(
                    "グラフ表示",
                    "グラフ化するノートがありません。\n(対象データが0件です)",
                    parent=self
                )
            return

        # 2. パフォーマンス制限
        if len(df) > 500:
            if not output_path:
                messagebox.showwarning(
                    "グラフ表示",
                    f"検索結果が多すぎます ({len(df)}件)。\nグラフ表示は500件に制限されます。",
                    parent=self
                )
            df = df.head(500)

        # 3. グラフ構築 (NetworkX)
        G = nx.DiGraph()
        notes_in_graph = set(df['key'])
        link_pattern = re.compile(r"\[\[(.*?)\]\]")

        # 3a. ノードを追加
        for index, row in df.iterrows():
            key = row.get('key')
            title = row.get('title', 'N/A')
            cp_key = row.get('commonplace_key', '').lower()

            icon_code = self.key_icons.get(cp_key, '•')
            color_hex = self.key_colors.get(cp_key, '#FFFFFF')

            # PDFへのURIを取得
            file_uri = get_pdf_uri_for_note(
                row, self.loaded_db_path, self.pdf_root_folder
            )

            tooltip = f"Key: {key}\nIndex: {cp_key}"
            if file_uri:
                tooltip += "\n(ダブルクリックしてPDFを開く)\n(右クリックでKeyをコピー)"

            G.add_node(
                key,
                label=title,
                title=tooltip,
                shape='icon',
                icon={
                    'code': icon_code,
                    'color': color_hex,
                    'size': 40
                },
                color=color_hex,
                # URIをノード属性に追加 (JavaScriptから参照)
                pdf_url=file_uri if file_uri else ""
            )

        # 3b. エッジを追加
        edge_count = 0
        for index, row in df.iterrows():
            source_key = row.get('key')
            memo = row.get('memo', '')

            for match in link_pattern.finditer(memo):
                full_match_content = match.group(1).strip()
                target_key = full_match_content.split(':')[0].strip()

                if target_key in notes_in_graph:
                    if source_key != target_key:
                        G.add_edge(source_key, target_key)
                        edge_count += 1

        print(f"グラフを生成: {len(df)} ノード, {edge_count} エッジ")

        # 4. 視覚化 (Pyvis)
        nt = Network(
            height="95vh",
            width="100%",
            bgcolor="#222222",
            font_color="white",
            directed=True,
            notebook=False
        )
        nt.from_nx(G)

        # 5. 物理演算のオプションを設定
        nt.set_options("""
        var options = {
          "physics": {
            "solver": "barnesHut",
            "barnesHut": {
              "gravitationalConstant": -8000,
              "centralGravity": 0.3,
              "springLength": 95,
              "springConstant": 0.04,
              "damping": 0.09
            },
            "minVelocity": 0.75
          },
          "interaction": {
            "tooltipDelay": 200,
            "hideEdgesOnDrag": true,
            "hover": true,
            "hoverConnectedEdges": true,
            "selectConnectedEdges": true,
            "navigationButtons": true,
            "keyboard": { "enabled": true }
          },
          "edges": {
            "arrows": {
              "to": { "enabled": true, "scaleFactor": 0.5 }
            },
            "color": {
              "color": "#848484",
              "highlight": "#FFFFFF",
              "hover": "#DDDDDD",
              "inherit": false
            },
            "smooth": {
              "type": "continuous",
              "forceDirection": "none",
              "roundness": 0.5
            }
          }
        }
        """)

        # 6. HTMLファイルとして保存し、ブラウザで開く
        try:
            if output_path:
                graph_file_path = output_path
            else:
                # デフォルトのパス (アプリと同じ場所)
                if getattr(sys, 'frozen', False):
                    base_path = Path(sys.executable).parent
                else:
                    base_path = Path(__file__).parent.parent
                graph_file_path = base_path / "synapsen_graph.html"

            # 6a. まずHTMLファイルを保存
            nt.save_graph(str(graph_file_path))

            # 6b. JavaScriptを拡張 (ダブルクリックでPDF + 右クリックでKeyコピー)
            custom_interaction_js = dedent("""
            document.addEventListener('DOMContentLoaded', function() {
                if (typeof network !== 'undefined') {

                    // --- ダブルクリック: PDFを開く ---
                    network.on("doubleClick", function(properties) {
                        var { nodes } = properties;
                        if (nodes.length > 0) {
                            var nodeId = nodes[0];
                            var nodeData = this.body.nodes[nodeId].options;
                            if (nodeData.pdf_url && nodeData.pdf_url !== "") {
                                window.open(nodeData.pdf_url, '_blank');
                            }
                        }
                    });

                    // --- 右クリック: Keyをコピー ---
                    network.on("oncontext", function (params) {
                        params.event.preventDefault();
                        var nodeId = network.getNodeAt(params.pointer.DOM);

                        if (nodeId) {
                            // ノードID (= Key) を取得
                            var textToCopy = nodeId;

                            // クリップボードへのコピーを試行
                            // (file:// プロトコル等での制限対策)
                            if (
                                navigator.clipboard &&
                                window.isSecureContext
                            ) {
                                navigator.clipboard.writeText(textToCopy).then(
                                    function() {
                                        // 成功時のフィードバック
                                        alert('Keyをコピーしました: ' + textToCopy);
                                    },
                                    function(err) {
                                        // 失敗時 -> プロンプトを表示
                                        prompt(
                                            "コピーしてください (Ctrl+C):",
                                            textToCopy
                                        );
                                    }
                                );
                            } else {
                                // API非対応環境用 -> プロンプトを表示
                                prompt(
                                    "コピーしてください (Ctrl+C):",
                                    textToCopy
                                );
                            }
                        }
                    });
                }
            });
            """)

            # 6c. 保存したHTMLを読み込む
            with open(graph_file_path, 'r', encoding='utf-8') as f:
                html_content = f.read()

            # 6d. </head> タグの直前に <script> ブロックを挿入
            script_tag = (
                "<script type=\"text/javascript\">\n" +
                f"{custom_interaction_js}\n</script>\n</head>"
            )
            html_content = html_content.replace("</head>", script_tag, 1)

            # 6e. 変更したHTMLを上書き保存
            with open(graph_file_path, 'w', encoding='utf-8') as f:
                f.write(html_content)

            # 6f. 変更後のHTMLをブラウザで開く
            if not output_path:
                webbrowser.open(graph_file_path.as_uri())

        except Exception as e:
            print(f"Graph display error: {e}")
            if not output_path:
                messagebox.showerror(
                    "グラフ表示エラー", f"グラフの生成または表示に失敗しました:\n{e}", parent=self)

    def show_local_graph(self):
        """
        現在詳細ペインで選択されているノートと、
        そのノートに「リンクしている」または「リンクされている」ノートのみでグラフを表示する。
        """
        if self.current_selected_row is None:
            messagebox.showinfo("情報", "ローカルグラフを表示するノートを選択してください。")
            return

        if self.df is None:
            return

        center_key = self.current_selected_row.get('key')
        if not center_key:
            return

        # 1. 関連するキーのセットを作成 (中心ノード + リンク先 + リンク元)
        related_keys = set()
        related_keys.add(center_key)

        # A. このノートが引用している先 (Forward Links)
        # memo欄から [[key]] を抽出して追加
        memo = self.current_selected_row.get('memo', '')
        link_pattern = re.compile(r"\[\[(.*?)\]\]")
        for match in link_pattern.finditer(memo):
            # [[key: title]] の形式も考慮し、:の前だけを取得
            content = match.group(1).split(':')[0].strip()
            related_keys.add(content)

        # B. このノートを引用している元 (Backlinks)
        # 全件走査で center_key を含んでいるノートを探す
        escaped_key = re.escape(center_key)
        # [[key]] または [[key:title]] にマッチ
        pattern = f"\\[\\[{escaped_key}[:\\]]"

        # 高速化のため、memo列が空でない行のみ対象にする等の工夫も可能ですが、
        # ここではシンプルに str.contains でフィルタリングします
        backlinks = self.df[
            self.df['memo'].str.contains(pattern, regex=True, na=False)]
        related_keys.update(backlinks['key'].tolist())

        # 2. 関連キーのみを含むDataFrameを作成
        local_df = self.df[self.df['key'].isin(related_keys)]

        if local_df.empty:
            messagebox.showinfo("情報", "関連するノートが見つかりませんでした。")
            return

        # 3. グラフ生成 (target_df を指定して呼び出し)
        print(f"Local Graph: {len(local_df)} notes related to {center_key}")
        self.generate_and_show_graph(target_df=local_df)

    # --- エクスポート機能 ---
    def export_search_results(self, include_pdf=False):
        """
        検索結果(または選択中)のデータをエクスポートする。
        include_pdf=True の場合、同じフォルダに統合PDFも生成する。
        """
        target_df = None
        export_mode = "search_results"  # メッセージ用

        # 1. エクスポート対象の決定
        if self.selected_keys:
            # 選択されている場合、そのノートのみを対象にする
            if self.df is not None:
                target_df = self.df[self.df['key'].isin(self.selected_keys)]
                export_mode = "selected_items"
        else:
            # 選択されていない場合、検索結果全体(キャッシュ)を使用
            target_df = self.filtered_df_cache

        if target_df is None or target_df.empty:
            messagebox.showinfo(
                "エクスポート",
                "エクスポートするノートがありません。\n(検索結果または選択アイテムが0件です)",
                parent=self
            )
            return

        # 2. 保存先フォルダをユーザーに選択させる
        export_folder_path = filedialog.askdirectory(title="エクスポート先フォルダを選択")
        if not export_folder_path:
            return

        export_path = Path(export_folder_path)
        now = datetime.datetime.now()
        timestamp = now.strftime("%Y%m%d_%H%M%S")

        if export_mode == "selected_items":
            folder_suffix = "Selected"
        else:
            folder_suffix = "Search"

        # ディレクトリ名を先に生成してから結合
        dir_name = f"Synapsen_Export_{folder_suffix}_{timestamp}"
        final_export_dir = export_path / dir_name

        try:
            final_export_dir.mkdir(parents=True, exist_ok=True)

            # 3. メタデータ保存 (target_df の長さを使用)
            meta_path = final_export_dir / "export_meta.txt"
            current_query = self.search_entry.get()

            with open(meta_path, 'w', encoding='utf-8') as f:
                # モード文字列を先に生成して行長を抑える
                mode_str = (
                    '選択アイテムのみ' if export_mode == 'selected_items' else '検索結果全体'
                )

                f.write("Synapsen Nexus エクスポート\n")
                f.write(f"モード: {mode_str}\n")
                f.write(f"件数: {len(target_df)} 件\n")
                f.write("=" * 30 + "\n")
                f.write(f"検索クエリ:\n{current_query}\n")

            # 4. CSV保存
            # CSVファイル名を変更
            csv_path = final_export_dir / "metadata.csv"

            # .drop() を使って full_text 列を明示的に削除
            df_to_export = target_df.drop(
                columns=['full_text'], errors='ignore'
            )

            # ( tags 列がリストの場合 ; 区切りに戻す - Nexus内では文字列のはずだが念のため )
            if 'tags' in df_to_export.columns:
                df_to_export['tags'] = df_to_export['tags'].apply(
                    lambda x: ";".join(x) if isinstance(x, list) else str(x)
                )

            df_to_export.to_csv(csv_path, index=False, encoding='utf-8-sig')

            # full_text を個別の .txt ファイルとして保存
            text_export_dir = final_export_dir / "FullText_Contents"
            text_export_dir.mkdir(exist_ok=True)

            print(f"FullText を {text_export_dir} にエクスポート中...")

            # 5. 本文テキスト保存
            # 元の (full_textを含む) DataFrame を使用
            for index, row in target_df.iterrows():
                key = row.get('key')
                text_content = row.get('full_text', '')

                if not key:
                    # keyが無いノートはスキップ (ほぼあり得ないが念のため)
                    continue

                # ファイル名は {key}.txt (例: 20240101000000.txt)
                text_file_path = text_export_dir / f"{key}.txt"

                with open(text_file_path, 'w', encoding='utf-8') as f:
                    f.write(text_content)

            # 6. グラフ (HTML) を保存
            graph_html_path = final_export_dir / "relation_graph.html"
            self.generate_and_show_graph(
                output_path=graph_html_path, target_df=target_df)

            if include_pdf:
                pdf_save_path = final_export_dir / "Merged_Notes.pdf"
                # 共通のPDF結合処理を呼び出す
                merge_result = self._execute_pdf_merge(
                    target_df, pdf_save_path)

                if not merge_result:
                    # 失敗時
                    messagebox.showwarning(
                        "PDF結合警告", "PDFの結合に失敗したか、対象ファイルがありませんでした。",
                        parent=self
                    )

            # 完了メッセージ
            msg = f"エクスポートが完了しました。\n保存先: {final_export_dir}"
            if include_pdf:
                msg += "\n(統合PDFも含みます)"

            messagebox.showinfo("完了", msg, parent=self)

        except Exception as e:
            print(f"エクスポート処理中にエラー: {e}")
            messagebox.showerror("エクスポートエラー", f"失敗しました:\n{e}", parent=self)

            messagebox.showinfo(
                "エクスポート完了",
                f"検索結果を指定のフォルダに保存しました:\n\n"
                f"・メタデータCSV (search_results_metadata.csv)\n"
                f"・本文テキスト (FullText_Contents フォルダ)\n"
                f"・関連グラフ (search_graph.html)\n\n"
                f"保存先:\n{final_export_dir}",
                parent=self
            )

        except Exception as e:
            print(f"エクスポート処理中にエラー: {e}")
            messagebox.showerror(
                "エクスポートエラー",
                f"エクスポートに失敗しました:\n{e}", parent=self
                )
            # (もしエラーが発生したら、中途半端に作成したフォルダを削除する)
            if final_export_dir.exists():
                try:
                    shutil.rmtree(final_export_dir)
                except Exception as e_del:
                    print(f"エラー後のエクスポートフォルダ削除に失敗: {e_del}")

    def merge_and_export_pdf(self):
        """メニューから「統合PDF」単体を選んだ場合のラッパー"""
        # 1. 対象データの決定
        target_df = None
        if self.selected_keys:
            if self.df is not None:
                target_df = self.df[self.df['key'].isin(self.selected_keys)]
        else:
            target_df = self.filtered_df_cache

        if target_df is None or target_df.empty:
            messagebox.showinfo("PDF結合", "出力対象のノートがありません。", parent=self)
            return

        # 2. 保存先ダイアログ (単体の場合はファイル保存ダイアログ)
        time_text = datetime.datetime.now().strftime('%Y%m%d')
        save_path = filedialog.asksaveasfilename(
            title="統合PDFを保存",
            defaultextension=".pdf",
            filetypes=[("PDF Files", "*.pdf")],
            initialfile=f"Merged_Notes_{time_text}.pdf"
        )
        if not save_path:
            return

        # 3. 実行
        if self._execute_pdf_merge(target_df, Path(save_path)):
            messagebox.showinfo("完了", f"PDFを保存しました:\n{save_path}", parent=self)
        else:
            messagebox.showwarning(
                "失敗", "結合できるPDFファイルが見つかりませんでした。", parent=self)

    def _execute_pdf_merge(self, target_df, save_path: Path):
        """
        指定されたDataFrameのノートをPDF結合し、save_pathに保存する。
        同時に「しおり(Bookmark)」を追加する。
        """
        try:
            writer = PdfWriter()
            processed_count = 0
            current_page_index = 0  # 結合後の現在のページ番号

            # 待機カーソル
            self.configure(cursor="watch")
            self.update()

            # データは日付順(またはリスト順)で処理する
            for index, row in target_df.iterrows():
                filepath_str = row.get('filepath', '')
                if not filepath_str:
                    continue

                # パス解決
                pdf_path = Path(filepath_str)
                if not pdf_path.is_absolute() and self.pdf_root_folder:
                    pdf_path = self.pdf_root_folder / pdf_path

                if pdf_path.is_file():
                    try:
                        reader = PdfReader(pdf_path)

                        # --- [しおり追加] ---
                        # 各ノートの先頭ページに、そのノートのタイトルでしおりを追加
                        # フォーマット: "YYYY/MM/DD タイトル"
                        date_str = row.get('date', '??????')
                        if len(date_str) == 8:
                            yyyy = date_str[:4]
                            mm = date_str[4:6]
                            dd = date_str[6:]
                            date_fmt = f"{yyyy}/{mm}/{dd}"
                        else:
                            date_fmt = date_str

                        title = row.get('title', 'No Title')
                        bookmark_title = f"{date_fmt} {title}"

                        # しおりを追加 (現在のページ位置を指定)
                        writer.add_outline_item(
                            title=bookmark_title,
                            page_number=current_page_index
                        )

                        # ページ追加
                        for page in reader.pages:
                            writer.add_page(page)
                            current_page_index += 1  # ページ番号を進める

                        processed_count += 1
                    except Exception as e:
                        print(f"PDF merge error ({pdf_path.name}): {e}")
                else:
                    print(f"File not found: {pdf_path}")

            if processed_count > 0:
                with open(save_path, "wb") as f:
                    writer.write(f)
                return True
            else:
                return False

        except Exception as e:
            print(f"PDF merge execution error: {e}")
            raise e
        finally:
            self.configure(cursor="")

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

        if self.df is None:
            messagebox.showwarning("データなし", "データベースが読み込まれていません。")
            return

        # 編集ウィンドウ (書き込み可能) のインスタンスを作成
        editor_win = NoteEditorWindow(
            self,
            note_data,
            self.commonplace_keys_options,
            self.predefined_tags,
            self.save_edit_callback
        )
        editor_win.focus()  # ウィンドウにフォーカスを当てる

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

        # 1. utilsのDB更新関数を呼び出す
        update_note_in_db(self.loaded_db_path, key_to_update, new_data_dict)

        # 2. 変更をUIに反映するため、DBを再読み込み
        print(f"ノート {key_to_update} を更新しました。DBを再読み込みします。")

        # 再表示したいキーを引数に渡す
        self.load_db_from_path(
            self.loaded_db_path, key_to_redisplay=key_to_update)

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
            parent=self
        )

        if answer:  # Yesが押されたら
            try:
                # 1. utilsのDB削除関数を呼び出す
                delete_note_from_db(self.loaded_db_path, key_to_delete)

                # 2. 変更をUIに反映するため、DBを再読み込み
                print(f"ノート {key_to_delete} を削除しました。DBを再読み込みします。")
                self.load_db_from_path(self.loaded_db_path)

                messagebox.showinfo(
                    "削除完了", f"ノート {key_to_delete} を削除しました。", parent=self)

            except Exception as e:
                messagebox.showerror(
                    "削除エラー", f"データベースからの削除に失敗しました:\n{e}", parent=self)


if __name__ == "__main__":
    app = Synapsen_Nexus()
    if app.icon_path:
        try:
            # 'default=' を指定し、OSダイアログ(エクスプローラ等)にも適用
            app.iconbitmap(default=str(app.icon_path))
        except Exception as e:
            print(f"Icon default setting error: {e}")
    else:
        print("警告: アイコンファイル (assets/synapsen.ico) が見つかりません。")
    app.mainloop()

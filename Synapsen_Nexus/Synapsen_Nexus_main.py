import customtkinter as ctk
from tkinter import filedialog, messagebox
import pandas as pd
from pathlib import Path
import re
import sys

# 分割したモジュールをインポート
from utils import (
    load_app_config, load_sql_data_file, open_pdf_viewer,
    build_memo_display, build_references_display, find_backlinks_df,
    update_note_in_db, delete_note_from_db
)
from search_parser import parse_or_expression

from preview_window import NotePreviewWindow
from editor_window import NoteEditorWindow


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
        self.geometry("1200x800")
        self.grid_columnconfigure(0, weight=3)
        self.grid_columnconfigure(1, weight=2)
        self.grid_rowconfigure(1, weight=1)

        # --- アプリケーションの状態変数 ---
        self.df = None  # ノートデータを保持するDataFrame
        self.pdf_root_folder = None  # config.iniから読み込むPDFのルートパス
        self.key_icons = {}  # IndexKeyごとのアイコン
        self.key_colors = {}  # IndexKeyごとの色
        self.commonplace_keys_options = []  # IndexKeyの全オプション
        self.predefined_tags = []  # オートコンプリート用のタグリスト
        self.loaded_db_path = None  # 現在開いているDBのパス
        self.filter_checkboxes = {}  # IndexKeyフィルターのチェックボックス変数
        self.filter_panel_expanded = False  # フィルターパネルが開いているか

        # --- オートコンプリート関連 ---
        self.selected_suggestion_index = -1
        self.current_suggestions = []

        self.create_widgets()
        self.load_config()

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
        utils.load_app_config を使用する。
        """
        try:
            # 実行ファイルのパスを基準にconfig.iniを探す
            if getattr(sys, 'frozen', False):
                base_path = Path(sys.executable).parent
            else:
                base_path = Path(__file__).parent

            # utilsから設定を辞書として読み込む
            config_data = load_app_config(base_path)

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
        """アプリケーションのUIコンポーネントを作成し、配置する。"""

        # --- トップフレーム (ファイル読み込みボタンと検索バー) ---
        top_frame = ctk.CTkFrame(self)
        top_frame.grid(
            row=0, column=0, columnspan=2, padx=10, pady=(10, 0), sticky="ew"
            )
        top_frame.grid_columnconfigure(1, weight=1)

        ctk.CTkButton(
            top_frame, text="目次データベースを開く", command=self.load_database_dialog
        ).pack(side="left", padx=5)

        search_container = ctk.CTkFrame(top_frame, fg_color="transparent")
        search_container.pack(side="left", fill="x", expand=True, padx=5)

        self.search_entry = ctk.CTkEntry(
            search_container,
            placeholder_text={
                "検索 (AND, OR, - , ( ) を使用可, プレフィックスを使用する事で検索対象を絞る" +
                "(例: date:YYYYMM / date:YYYYMMDD))"
                }
        )
        self.search_entry.pack(fill="x")

        # 本文検索(FTS)をトグルするチェックボックス
        self.fts_checkbox = ctk.CTkCheckBox(
            top_frame, text="本文検索"
        )
        self.fts_checkbox.pack(side="left", padx=5)

        # チェックボックスの状態が変わったら検索を再実行
        self.fts_checkbox.configure(command=self.perform_search)

        # 検索バーのイベントバインド
        self.search_entry.bind("<KeyRelease>", self.handle_keyrelease)
        self.search_entry.bind("<FocusOut>", self.hide_autocomplete)
        self.search_entry.bind("<FocusIn>", self.update_suggestions)
        self.search_entry.bind("<Down>", self.navigate_suggestions)
        self.search_entry.bind("<Up>", self.navigate_suggestions)
        self.search_entry.bind("<Return>", self.confirm_suggestion)

        # オートコンプリート用の非表示フレーム
        self.autocomplete_frame = ctk.CTkScrollableFrame(self, label_text="")

        # --- 左パネル (フィルターと検索結果) ---
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

        # --- 右パネル (詳細表示) ---
        self.details_frame = ctk.CTkFrame(self)
        self.details_frame.grid(
            row=1, column=1, padx=(0, 10), pady=10, sticky="nsew"
            )

        self.details_frame.grid_rowconfigure(5, weight=2)  # <--- メモ欄 (重み2)
        self.details_frame.grid_rowconfigure(7, weight=1)  # <--- 引用元欄 (重み1)
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

        ctk.CTkLabel(
            self.details_frame, text="メモ:", anchor="w"
            ).grid(row=4, column=0, padx=10, pady=5, sticky="nw")

        # メモ表示用 (utils.build_memo_display で中身が構築される)
        self.memo_display_frame = ctk.CTkScrollableFrame(self.details_frame)
        self.memo_display_frame.grid(
            row=5, column=1, padx=10, pady=5, sticky="nsew"
            )

        # 引用元欄
        ctk.CTkLabel(
            self.details_frame, text="引用元:", anchor="w"
            ).grid(row=6, column=0, padx=10, pady=5, sticky="nw")

        self.references_display_frame = ctk.CTkScrollableFrame(
            self.details_frame, label_text="このノートを引用しているノート"
            )
        self.references_display_frame.grid(
            row=7, column=1, padx=10, pady=5, sticky="nsew"
            )

        self.edit_button_frame = ctk.CTkFrame(
            self.details_frame, fg_color="transparent")
        self.edit_button_frame.grid(row=8, column=0, columnspan=2, pady=10)

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

    # --- オートコンプリート関連メソッド ---

    def handle_keyrelease(self, event):
        """検索バーでのキー入力（リリース）イベントを処理する。"""
        if event.keysym in ("Up", "Down", "Return", "Escape"):
            return
        self.update_suggestions()
        self.perform_search()

    def update_suggestions(self, event=None):
        """検索バーの入力に基づき、オートコンプリートの候補を更新する。"""
        self.selected_suggestion_index = -1
        query = self.search_entry.get()
        # " AND " や " OR " で区切られた最後の単語を取得
        last_word = re.split(
            r'\s+(?:AND|OR)\s+', query, flags=re.IGNORECASE)[-1].strip()

        suggestions = []
        if query == "" or query.upper().endswith(" AND ") or query.upper().endswith(" OR "):
            # オペレータの後は全タグリストを表示
            suggestions = self.predefined_tags
        elif last_word:
            # 入力中の単語で前方一致検索
            suggestions = [tag for tag in self.predefined_tags if tag.lower().startswith(last_word.lower())]

        if suggestions:
            self.show_autocomplete(suggestions)
        else:
            self.hide_autocomplete()

    def show_autocomplete(self, suggestions):
        """オートコンプリートの候補リストウィンドウを表示する。"""
        self.current_suggestions = suggestions
        for widget in self.autocomplete_frame.winfo_children():
            widget.destroy()

        for i, suggestion in enumerate(suggestions):
            fg_color = "gray30" if i == self.selected_suggestion_index else "transparent"
            btn = ctk.CTkButton(
                self.autocomplete_frame, text=suggestion, fg_color=fg_color,
                text_color=ctk.ThemeManager.theme["CTkLabel"]["text_color"],
                anchor="w", command=lambda s=suggestion: self.select_suggestion(s)
            )
            btn.pack(fill="x", padx=5, pady=2)

        # 検索バーの真下に配置
        x = self.search_entry.winfo_rootx() - self.winfo_rootx()
        y = self.search_entry.winfo_rooty() - self.winfo_rooty() + self.search_entry.winfo_height()
        width = self.search_entry.winfo_width()
        height = min(200, len(suggestions) * 35)

        self.autocomplete_frame.configure(width=width, height=height)
        self.autocomplete_frame.place(x=x, y=y)
        self.autocomplete_frame.lift()

    def select_suggestion(self, suggestion):
        """オートコンプリート候補をクリックまたはEnterで選択したときの処理。"""
        query = self.search_entry.get()

        # 現在入力中の単語を、選択した候補で置き換える
        match = re.search(
            r'(\s+(?:AND|OR)\s+)?([^\s,]*)$', query, re.IGNORECASE
            )
        if match:
            preceding_operator = match.group(1) if match.group(1) else ''
            base_query = query[:match.start()]
            new_query = f"{base_query}{preceding_operator}{suggestion} "
        else:
            new_query = f"{suggestion} "

        self.search_entry.delete(0, "end")
        self.search_entry.insert(0, new_query)
        self.search_entry.focus_force()
        self.search_entry.icursor("end")

        self.hide_autocomplete()
        self.perform_search()

    def hide_autocomplete(self, event=None):
        """オートコンプリートウィンドウを非表示にする。"""
        # 少し遅延させて非表示にし、クリックイベントが発火できるようにする
        self.after(200, lambda: self.autocomplete_frame.place_forget())

    def navigate_suggestions(self, event):
        """キーボードの上下矢印キーで候補リストを移動する。"""
        if not self.autocomplete_frame.winfo_ismapped() or not self.current_suggestions:
            return

        num_suggestions = len(self.current_suggestions)
        if event.keysym == "Down":
            self.selected_suggestion_index = (self.selected_suggestion_index + 1) % num_suggestions
        elif event.keysym == "Up":
            self.selected_suggestion_index = (self.selected_suggestion_index - 1 + num_suggestions) % num_suggestions

        # 選択項目がリストに表示されるようにスクロール
        self.autocomplete_frame._parent_canvas.yview_moveto(
            self.selected_suggestion_index / num_suggestions
        )
        # 選択ハイライトを更新
        self.show_autocomplete(self.current_suggestions)
        return "break"  # 他のキーバインドを抑制

    def confirm_suggestion(self, event):
        """Enterキーで選択中の候補を確定する。"""
        if self.autocomplete_frame.winfo_ismapped() and self.selected_suggestion_index != -1:
            self.select_suggestion(
                self.current_suggestions[self.selected_suggestion_index]
            )
            return "break"  # 検索が二重に実行されるのを防ぐ

        # 候補が選択されていない場合は、通常の検索を実行
        self.perform_search()
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
            selected_keys = [key for key, var in self.filter_checkboxes.items() if var.get() == '1']
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

    def load_db_from_path(self, filepath: Path):
        """
        指定されたパスからDBを読み込み、DataFrameを更新する。
        utils.load_sql_data_file を使用する。

        Args:
            filepath (Path): 読み込むDBファイルのパス。
        """
        try:
            # utilsの関数でDataFrameを読み込む
            self.df = load_sql_data_file(filepath)
            self.loaded_db_path = filepath

            # UIをリセット・更新
            self.perform_search()
            self.clear_details()
            self.filter_panel_expanded = False
            self.sync_filter_panel_view()

        except Exception as e:
            messagebox.showerror("データベース読み込みエラー", str(e))

    def populate_key_filters(self):
        """config.iniの情報に基づき、IndexKeyフィルターのUIを構築する。"""
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
                command=self.perform_search  # チェック時に検索を再実行
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
        selected_keys = [key for key, var in self.filter_checkboxes.items() if var.get() == '1']
        if selected_keys:
            filtered_df = filtered_df[filtered_df['commonplace_key'].isin(selected_keys)]

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

        self.update_results_list(filtered_df)
        self.update_collapsed_filter_view()

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
            # シングルクリックで詳細表示
            command = lambda e, r=row: self.show_details(r)
            item_frame.bind("<Button-1>", command)
            icon_label.bind("<Button-1>", command)
            text_label.bind("<Button-1>", command)

            # ダブルクリックでPDFを開く
            pdf_command = lambda e, r=row: self.open_pdf(r)
            item_frame.bind("<Double-Button-1>", pdf_command)
            icon_label.bind("<Double-Button-1>", pdf_command)
            text_label.bind("<Double-Button-1>", pdf_command)

    def clear_details(self):
        """詳細表示ペインの内容をすべてクリアする。"""
        self.title_label.configure(text="")
        self.key_label.configure(text="")
        self.cpkey_label.configure(text="")
        self.tags_label.configure(text="")

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

        row = row_data  # 分かりやすくするため
        self.title_label.configure(text=row.get('title', ''))
        self.key_label.configure(text=row.get('key', ''))
        self.cpkey_label.configure(text=row.get('commonplace_key', ''))

        # タグ表示（文字列をリストに変換して表示）
        tags_str = str(row.get('tags', ''))
        tags_list = [tag for tag in tags_str.split(';') if tag]
        self.tags_label.configure(text=", ".join(tags_list))

        # メモ表示（リンク構築）
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
        current_key = row.get('key', '')

        # utilsの新関数を使って引用元DFを取得
        backlinks_df = find_backlinks_df(self.df, current_key)

        # utilsの新関数を使って引用元UIを構築
        build_references_display(
            self.references_display_frame,
            backlinks_df,
            self.open_preview_window,  # <-- リンククリック時のコールバック
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
            self.commonplace_key_options,
            self.predefined_tags,
            self.save_edit_callback  # <-- 保存時に呼んでほしい関数
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
        self.load_db_from_path(self.loaded_db_path)

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
    if app.icon_path:  # <-- クラス内で取得したパスを利用
        try:
            # 'default=' を指定し、OSダイアログ(エクスプローラ等)にも適用
            app.iconbitmap(default=str(app.icon_path))
        except Exception as e:
            print(f"Icon default setting error: {e}")
    else:
        print("警告: アイコンファイル (assets/synapsen.ico) が見つかりません。")
    app.mainloop()

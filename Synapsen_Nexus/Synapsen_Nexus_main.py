import customtkinter as ctk
from tkinter import filedialog, messagebox
import pandas as pd
from pathlib import Path
import re
import sys

# 分割したモジュールをインポート
from utils import (
    load_app_config, load_csv_data_file, open_pdf_viewer,
    build_memo_display, build_references_display, find_backlinks_df
)
from search_parser import parse_or_expression
from preview_window import NotePreviewWindow


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
        self.loaded_csv_path = None  # 現在開いているCSVのパス
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

            # デフォルトCSVが設定されていれば自動で読み込む
            default_csv_path = config_data.get('default_csv_path')
            if default_csv_path and default_csv_path.is_file():
                self.load_csv_from_path(default_csv_path)
                # print(f"[DEBUG] CSV: {default_csv_path}")
            else:
                if default_csv_path:
                    print(f"デフォルトCSVが見つかりません: {default_csv_path}")
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
            top_frame, text="目次CSVファイルを開く", command=self.load_csv_data
        ).pack(side="left", padx=5)

        search_container = ctk.CTkFrame(top_frame, fg_color="transparent")
        search_container.pack(side="left", fill="x", expand=True, padx=5)

        self.search_entry = ctk.CTkEntry(
            search_container,
            placeholder_text="検索 (AND, OR, - , ( ) を使用可, プレフィックスを使用する事で検索対象を絞る(例: date:YYYYMM / date:YYYYMMDD))"
        )
        self.search_entry.pack(fill="x")

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
        last_word = re.split(r'\s+(?:AND|OR)\s+', query, flags=re.IGNORECASE)[-1].strip()

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

    def load_csv_data(self):
        """「目次CSVファイルを開く」ボタンの動作。ファイルダイアログを開く。"""
        filepath = filedialog.askopenfilename(
            title="目次CSVファイルを選択",
            filetypes=[("CSV files", "*.csv")]
        )
        if not filepath:
            return
        self.load_csv_from_path(filepath)

    def load_csv_from_path(self, filepath):
        """
        指定されたパスからCSVを読み込み、DataFrameを更新する。
        utils.load_csv_data_file を使用する。

        Args:
            filepath (str or Path): 読み込むCSVファイルのパス。
        """
        try:
            # utilsの関数でDataFrameを読み込む
            self.df = load_csv_data_file(filepath)
            self.loaded_csv_path = filepath

            # UIをリセット・更新
            self.perform_search()
            self.clear_details()
            self.filter_panel_expanded = False
            self.sync_filter_panel_view()

        except Exception as e:
            messagebox.showerror("CSV読み込みエラー", str(e))

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
        if query_text:
            try:
                # search_parserの関数を呼び出し
                final_mask = parse_or_expression(filtered_df, query_text)
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

            icon_label = ctk.CTkLabel(item_frame, text=icon, text_color=color, font=("", 16), width=20)
            icon_label.pack(side="left")

            display_text = f"[{row.get('date')}] {row.get('title', 'N/A')}"
            text_label = ctk.CTkLabel(item_frame, text=display_text, anchor="w")
            text_label.pack(side="left", fill="x", expand=True)

            # --- イベントバインド ---
            # シングルクリックで詳細表示
            command = lambda e, idx=index: self.show_details(idx)
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
        指定されたキーのノートを新しいプレビューウィンドウで開く。

        Args:
            key (str): 表示するノートの 'key' (ID)。
        """
        if self.df is None:
            messagebox.showwarning("データなし", "CSVデータが読み込まれていません。")
            return

        target_note_row = self.df[self.df['key'] == key]

        if target_note_row.empty:
            messagebox.showwarning("ノート不明", f"ID '{key}' に一致するノートが見つかりませんでした。")
            return

        note_data = target_note_row.iloc[0]

        # プレビューウィンドウのインスタンスを作成
        preview_win = NotePreviewWindow(self, note_data)
        preview_win.focus()  # ウィンドウにフォーカスを当てる

    def show_details(self, index):
        """
        選択されたノートの詳細を右ペインに表示する。
        utils.build_memo_display を使用してメモ欄を構築する。

        Args:
            index (int): 表示するノートのDataFrameインデックス。
        """
        if self.df is None or index not in self.df.index:
            return

        row = self.df.loc[index]
        self.title_label.configure(text=row.get('title', ''))
        self.key_label.configure(text=row.get('key', ''))
        self.cpkey_label.configure(text=row.get('commonplace_key', ''))
        self.tags_label.configure(
            text=str(row.get('tags', '')).replace(';', ', ')
            )

        # utilsの関数でメモ欄を構築
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
            self.open_preview_window,  # リンククリック時のコールバック
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
            self.loaded_csv_path,
            self.pdf_root_folder
        )


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

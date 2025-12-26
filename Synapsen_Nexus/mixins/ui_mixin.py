import customtkinter as ctk
import logging
import sys
from pathlib import Path

logger = logging.getLogger(__name__)

# === 2. プロジェクトルートをパスに追加 ===
current_dir = Path(__file__).parent
root_dir = current_dir.parent
if str(root_dir) not in sys.path:
    sys.path.append(str(root_dir))

from theme import SemanticColors as Colors  # noqa: E402


class NexusUiMixin:
    """UI構築と更新ロジックを担当するMixin"""

    def create_widgets(self):
        """全ウィジェットの生成と配置"""
        # トップコンテナ
        top_container = ctk.CTkFrame(self, fg_color=Colors.BACKGROUND_PANEL)
        top_container.grid(
            row=0, column=0, columnspan=2, padx=10, pady=(10, 0), sticky="ew"
        )

        # --- 1段目: 検索バー等 ---
        row1 = ctk.CTkFrame(top_container, fg_color="transparent")
        row1.pack(side="top", fill="x", expand=True, pady=(5, 2), padx=5)

        self._create_left_buttons(row1)
        self._create_smart_search_buttons(row1)  # 右端を先に配置
        self._create_search_list_callup_buttons(row1)
        self._create_search_bar(row1)  # 残りを埋める

        # --- 2段目: ツールバー ---
        row2 = ctk.CTkFrame(top_container, fg_color="transparent")
        row2.pack(side="top", fill="x", pady=(0, 5), padx=5)

        self._create_view_tools(row2)
        self._create_action_tools(row2)
        self._create_extra_tools(row2)

        # --- パネル構成 ---
        # 左パネル
        self.left_panel = ctk.CTkFrame(self, fg_color="transparent")
        self.left_panel.grid(row=1, column=0, padx=10, pady=10, sticky="nsew")
        self.left_panel.grid_rowconfigure(2, weight=1)
        self.left_panel.grid_columnconfigure(0, weight=1)

        self._create_left_panel_contents()

        # 右パネル
        self._create_details_panel()

    # -------------------------------------------------------------------------
    # ヘルパーメソッド (ウィジェット生成)
    # -------------------------------------------------------------------------

    def _create_left_buttons(self, parent):
        frame = ctk.CTkFrame(parent, fg_color="transparent")
        frame.pack(side="left", padx=(0, 5))
        # load_database_dialog は Synapsen_Nexus_main.py に定義されている前提
        ctk.CTkButton(
            frame,
            text="DB",
            command=self.load_database_dialog,
            width=50,
            fg_color=Colors.UI_BASIC,
            hover_color=Colors.adjust_brightness(Colors.UI_BASIC),
            text_color=Colors.adjust_brightness(Colors.UI_BASIC, factor=0.2),
        ).pack(side="left", padx=(0, 5))
        # show_search_help はこのファイル(Mixin)の下部に定義
        ctk.CTkButton(
            frame,
            text="？",
            command=self.show_search_help,
            width=30,
            fg_color=Colors.UI_BASIC,
            hover_color=Colors.adjust_brightness(Colors.UI_BASIC),
            text_color=Colors.adjust_brightness(Colors.UI_BASIC, factor=0.2),
        ).pack(side="left", padx=0)

    def _create_smart_search_buttons(self, parent):
        """検索保存・呼び出しボタンの作成 (ここで saved_search_combo を定義)"""
        frame = ctk.CTkFrame(parent, fg_color="transparent")
        frame.pack(side="right", padx=(5, 0))

        # self.search_manager はメインクラスで初期化されている
        self.save_search_button = ctk.CTkButton(
            frame,
            text="検索保存",
            command=self.search_manager.save_current_search,
            width=80,
            fg_color=Colors.UI_BASIC,
            hover_color=Colors.adjust_brightness(Colors.UI_BASIC),
            text_color=Colors.adjust_brightness(Colors.UI_BASIC, factor=0.2),
        )
        self.save_search_button.pack(side="left", padx=(0, 5))

        self.saved_search_combo = ctk.CTkComboBox(
            frame,
            values=["保存済み検索..."],
            width=150,
            command=self.search_manager.on_saved_search_selected,
            button_color=Colors.adjust_brightness(Colors.UI_SETTING, 1.1),
            button_hover_color=Colors.UI_SETTING,
            dropdown_fg_color=(Colors.BACKGROUND_HOLLOW, Colors.BACKGROUND_DARK_HOLLOW),
            dropdown_hover_color=(
                Colors.adjust_brightness(Colors.BACKGROUND_HOLLOW, 0.85),
                Colors.adjust_brightness(Colors.BACKGROUND_DARK_HOLLOW, 0.15),
            ),
        )
        self.saved_search_combo.pack(side="left", padx=0)
        self.saved_search_combo.set("保存済み検索...")

    def _create_search_list_callup_buttons(self, parent: ctk.CTkFrame) -> None:
        """
        検索に使用するリストウィンドウを呼び出すボタンを生成する。
        """
        frame = ctk.CTkFrame(parent, fg_color="transparent")
        frame.pack(side="right", padx=(5, 0))

        # タグリスト
        self.callup_tag_list_button = ctk.CTkButton(
            frame,
            text="Tag",
            command=self.open_tag_window,
            width=30,
            fg_color=Colors.UI_BASIC,
            hover_color=Colors.adjust_brightness(Colors.UI_BASIC),
            text_color=Colors.adjust_brightness(Colors.UI_BASIC, factor=0.2),
        )
        self.callup_tag_list_button.pack(side="left", padx=0)

    def _create_search_bar(self, parent):
        search_container = ctk.CTkFrame(parent, fg_color="transparent")
        search_container.pack(side="left", fill="x", expand=True, padx=5)

        self.search_entry = ctk.CTkEntry(
            search_container, placeholder_text="検索 (例: tag:Idea ...)"
        )
        self.search_entry.pack(fill="x", expand=True)

        # イベントバインド
        self.search_entry.bind("<KeyRelease>", self.handle_keyrelease)

        # Enterキーは直接検索実行へバインド
        self.search_entry.bind("<Return>", lambda e: self._trigger_search_now())

    def _create_view_tools(self, parent):
        frame = ctk.CTkFrame(parent, fg_color="transparent")
        frame.pack(side="left", padx=0)

        self.sort_button = ctk.CTkButton(
            frame,
            text="▲ 古い順",
            command=self.toggle_sort_order,
            width=90,
            fg_color=Colors.UI_SECONDARY,
            hover_color=Colors.adjust_brightness(Colors.UI_SECONDARY),
        )
        self.sort_button.pack(side="left", padx=(0, 5))

        self.fts_checkbox = ctk.CTkCheckBox(
            frame,
            text="本文・メモ検索",
            command=self._trigger_search_now,
            fg_color=Colors.UI_BASIC,
            hover_color=Colors.adjust_brightness(Colors.UI_BASIC),
            checkmark_color=Colors.adjust_brightness(Colors.UI_BASIC, 1.8),
        )
        self.fts_checkbox.pack(side="left", padx=(0, 10))

        self.exclude_tags_checkbox = ctk.CTkCheckBox(
            frame,
            text="除外タグ",
            command=self._trigger_search_now,
            fg_color=Colors.UI_BASIC,
            hover_color=Colors.adjust_brightness(Colors.UI_BASIC),
            checkmark_color=Colors.adjust_brightness(Colors.UI_BASIC, 1.8),
        )
        self.exclude_tags_checkbox.pack(side="left", padx=(0, 10))
        defaults = getattr(self, "exclude_tags_by_default", [])
        if defaults:
            # タグ名を表示して分かりやすくする (例: "除外: Archive")
            label_text = f"除外: {','.join(defaults)}"
            if len(label_text) > 20:  # 長すぎる場合は省略
                label_text = "除外タグ適用"

            self.exclude_tags_checkbox.configure(text=label_text)
            self.exclude_tags_checkbox.select()  # 初期状態でONにする
        else:
            # 設定がない場合はチェックボックスを無効化
            self.exclude_tags_checkbox.configure(state="disabled", text="除外設定なし")

        self.selection_info_label = ctk.CTkLabel(
            frame, text="選択: 0", font=("", 12, "bold"), text_color="gray", width=60
        )
        self.selection_info_label.pack(side="left", padx=(0, 5))

        self.clear_selection_button = ctk.CTkButton(
            frame,
            text="×",
            command=self.clear_selection,
            width=30,
            fg_color=Colors.adjust_brightness(Colors.UI_CANCEL, 1.1),
            hover_color=Colors.UI_CANCEL,
            text_color="white",
        )
        self.clear_selection_button.pack(side="left", padx=0)

        ctk.CTkLabel(parent, text="|", text_color="gray").pack(side="left", padx=5)

    def _create_action_tools(self, parent):
        frame = ctk.CTkFrame(parent, fg_color="transparent")
        frame.pack(side="left", padx=0)

        self.graph_menu_var = ctk.StringVar(value="グラフ表示")
        self.graph_menu = ctk.CTkOptionMenu(
            frame,
            variable=self.graph_menu_var,
            values=["全体 (Global)", "関連 (Local)", "選択 (Selected)"],
            command=self.handle_graph_menu,
            width=130,
            fg_color=Colors.UI_SECONDARY,
            button_color=Colors.adjust_brightness(Colors.UI_SECONDARY),
            button_hover_color=Colors.adjust_brightness(Colors.UI_SECONDARY, 0.6),
            dropdown_fg_color=Colors.BACKGROUND_PANEL,
            dropdown_hover_color=Colors.adjust_brightness(Colors.BACKGROUND_PANEL),
        )
        self.graph_menu.pack(side="left", padx=(0, 5))

        self.copy_links_button = ctk.CTkButton(
            frame,
            text="リンクコピー",
            command=self.copy_selected_links,
            width=90,
            fg_color=Colors.UI_LINK,
            hover_color=Colors.adjust_brightness(Colors.UI_LINK),
            state="disabled",
        )
        self.copy_links_button.pack(side="left", padx=(0, 5))

        self.export_menu_var = ctk.StringVar(value="エクスポート")
        self.export_menu = ctk.CTkOptionMenu(
            frame,
            variable=self.export_menu_var,
            values=[
                "データ (CSV/TXT)",
                "統合PDF (Merge)",
                "全て (Data + PDF)",
                "MOC (Markdown)",
            ],
            command=self.handle_export_menu,
            width=130,
            fg_color=Colors.UI_EXPORT,
            button_color=Colors.adjust_brightness(Colors.UI_EXPORT),
            button_hover_color=Colors.adjust_brightness(Colors.UI_EXPORT, 0.4),
            dropdown_fg_color=Colors.BACKGROUND_PANEL,
            dropdown_hover_color=Colors.adjust_brightness(Colors.BACKGROUND_PANEL),
        )
        self.export_menu.pack(side="left", padx=(0, 5))
        ctk.CTkLabel(parent, text="|", text_color="gray").pack(side="left", padx=5)

    def _create_extra_tools(self, parent):
        frame = ctk.CTkFrame(parent, fg_color="transparent")
        frame.pack(side="left", padx=0)

        self.random_note_button = ctk.CTkButton(
            frame,
            text="閃き (R)",
            command=self.show_random_note,
            width=70,
            fg_color=Colors.UI_SECONDARY,
            hover_color=Colors.adjust_brightness(Colors.UI_SECONDARY),
        )
        self.random_note_button.pack(side="left", padx=(0, 5))

        self.canvas_button = ctk.CTkButton(
            frame,
            text="キャンバス",
            command=self.open_canvas,
            width=80,
            fg_color=Colors.CANVAS,
            hover_color=Colors.adjust_brightness(Colors.CANVAS),
            text_color=Colors.adjust_brightness(Colors.CANVAS, factor=0.2),
        )
        self.canvas_button.pack(side="left", padx=0)

        self.toggle_details_button = ctk.CTkButton(
            frame,
            text="▶ 詳細",
            command=self.toggle_details_panel,
            width=60,
            fg_color="transparent",
            border_width=1,
            hover_color=Colors.adjust_brightness(Colors.BACKGROUND_PANEL),
            text_color=("gray10", "gray90"),
        )
        self.toggle_details_button.pack(side="left", padx=(5, 0))

    def _create_left_panel_contents(self):
        # フィルタ、リスト、ページネーション
        filter_container = ctk.CTkFrame(
            self.left_panel, fg_color=Colors.BACKGROUND_PANEL
        )
        filter_container.grid(row=0, column=0, sticky="ew")
        filter_container.grid_columnconfigure(1, weight=1)

        self.toggle_filter_button = ctk.CTkButton(
            filter_container,
            text="▶ IndexKey",
            command=self.toggle_filter_panel,
            width=20,
            fg_color=Colors.UI_BASIC,
            hover_color=Colors.adjust_brightness(Colors.UI_BASIC),
            text_color=Colors.adjust_brightness(Colors.UI_BASIC, factor=0.2),
        )
        self.toggle_filter_button.grid(row=0, column=0, padx=5, pady=5)

        self.collapsed_icons_frame = ctk.CTkFrame(
            filter_container, fg_color="transparent"
        )
        self.collapsed_icons_frame.grid(row=0, column=1, padx=5, pady=5, sticky="w")

        self.key_filter_frame = ctk.CTkScrollableFrame(
            self.left_panel, label_text="", fg_color=Colors.BACKGROUND_PANEL
        )
        self.key_filter_frame.grid(row=1, column=0, padx=0, pady=(0, 5), sticky="nsew")
        self.key_filter_frame.grid_remove()  # 初期非表示

        self.results_list = ctk.CTkScrollableFrame(
            self.left_panel,
            label_text="ノート一覧",
            fg_color=Colors.BACKGROUND_PANEL,
            label_fg_color=Colors.adjust_brightness(Colors.BACKGROUND_PANEL, 0.8),
        )
        self.results_list.grid(row=2, column=0, padx=0, pady=0, sticky="nsew")

        # ページネーション
        self.pagination_frame = ctk.CTkFrame(
            self.left_panel, fg_color="transparent", height=40
        )
        self.pagination_frame.grid(row=3, column=0, sticky="ew", pady=(5, 0))

        self.btn_prev_page = ctk.CTkButton(
            self.pagination_frame,
            text="< 前",
            width=80,
            fg_color=Colors.UI_BASIC,
            hover_color=Colors.adjust_brightness(Colors.UI_BASIC),
            text_color=Colors.adjust_brightness(Colors.UI_BASIC, factor=0.2),
            command=self.prev_page,
            state="disabled",
        )
        self.btn_prev_page.pack(side="left", padx=10)
        self.page_label = ctk.CTkLabel(self.pagination_frame, text="0 / 0")
        self.page_label.pack(side="left", expand=True)
        self.btn_next_page = ctk.CTkButton(
            self.pagination_frame,
            text="次 >",
            width=80,
            fg_color=Colors.UI_BASIC,
            hover_color=Colors.adjust_brightness(Colors.UI_BASIC),
            text_color="black",
            command=self.next_page,
            state="disabled",
        )
        self.btn_next_page.pack(side="right", padx=10)

    def _create_details_panel(self):
        self.details_frame = ctk.CTkFrame(self, fg_color=Colors.BACKGROUND_PANEL)
        self.details_frame.grid(row=1, column=1, padx=(0, 10), pady=10, sticky="nsew")

        self.details_frame.grid_rowconfigure(0, weight=0)  # Info
        self.details_frame.grid_rowconfigure(1, weight=1)  # Memo
        self.details_frame.grid_rowconfigure(2, weight=1)  # Refs
        self.details_frame.grid_rowconfigure(3, weight=0)  # Buttons
        self.details_frame.grid_columnconfigure(0, weight=1)

        # トップ情報エリア
        top_info = ctk.CTkFrame(self.details_frame, fg_color="transparent")
        top_info.grid(row=0, column=0, sticky="ew", padx=5, pady=5)
        top_info.grid_columnconfigure(0, weight=1)
        top_info.grid_columnconfigure(1, weight=0)

        self.text_info_frame = ctk.CTkFrame(top_info, fg_color="transparent")
        self.text_info_frame.grid(row=0, column=0, sticky="nsew")

        self.preview_frame = ctk.CTkFrame(top_info, fg_color="transparent")
        self.preview_frame.grid(row=0, column=1, sticky="n", padx=(5, 0))
        self.pdf_preview_label = ctk.CTkLabel(
            self.preview_frame,
            text="No Preview",
            fg_color="gray20",
            width=200,
            height=280,
        )
        self.pdf_preview_label.pack()

        # メモと引用元のエリアのラベル色
        label_fg_color = (
            Colors.adjust_brightness(Colors.BACKGROUND_HOLLOW, 0.85),
            Colors.adjust_brightness(Colors.BACKGROUND_HOLLOW, 0.15),
        )

        # メモエリア
        self.memo_display_frame = ctk.CTkScrollableFrame(
            self.details_frame,
            label_text="メモ",
            height=150,
            fg_color=(Colors.BACKGROUND_HOLLOW, Colors.BACKGROUND_DARK_HOLLOW),
            label_fg_color=label_fg_color,
        )
        self.memo_display_frame.grid(row=2, column=0, sticky="nsew", padx=10)

        # 引用エリア
        self.references_display_frame = ctk.CTkScrollableFrame(
            self.details_frame,
            label_text="引用元",
            height=100,
            fg_color=(Colors.BACKGROUND_HOLLOW, Colors.BACKGROUND_DARK_HOLLOW),
            label_fg_color=label_fg_color,
        )
        self.references_display_frame.grid(
            row=3, column=0, sticky="nsew", padx=10, pady=5
        )

        # ボタンエリア
        self.edit_button_frame = ctk.CTkFrame(
            self.details_frame, fg_color="transparent"
        )
        self.edit_button_frame.grid(row=4, column=0, pady=10)

        self.open_preview_button = ctk.CTkButton(
            self.edit_button_frame,
            text="詳細プレビュー",
            command=self.open_current_note_in_preview,
            state="disabled",
            fg_color=Colors.UI_PREVIEW,
            hover_color=Colors.adjust_brightness(Colors.UI_PREVIEW),
        )
        self.open_preview_button.pack(side="left", padx=5)
        self.edit_button = ctk.CTkButton(
            self.edit_button_frame,
            text="編集",
            command=self.open_edit_dialog,
            state="disabled",
            fg_color=Colors.UI_EDIT,
            hover_color=Colors.adjust_brightness(Colors.UI_EDIT, 0.6),
        )
        self.edit_button.pack(side="left", padx=5)
        self.delete_button = ctk.CTkButton(
            self.edit_button_frame,
            text="削除",
            command=self.confirm_delete_note,
            fg_color=Colors.LABEL_DENGER,
            hover_color=Colors.adjust_brightness(Colors.LABEL_DENGER, 0.6),
            state="disabled",
        )
        self.delete_button.pack(side="left", padx=5)

    # -------------------------------------------------------------------------
    # UI更新メソッド
    # -------------------------------------------------------------------------

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
        total = getattr(self, "total_items", 0)
        current_page = getattr(self, "current_page", 0)

        if total > 0:
            label = f"検索結果: {total} 件 (ページ {current_page + 1})"
        else:
            label = f"検索結果 ({len(rows)}件)"

        self.results_list.configure(label_text=label)

        for row in rows:
            # 行フレームの作成
            item_frame = ctk.CTkFrame(self.results_list, fg_color="transparent")
            item_frame.pack(fill="x", padx=5, pady=2)

            # --- チェックボックス ---
            note_key = row.get("key")
            # selected_keys は Mainクラスで管理
            is_selected = note_key in self.selected_keys

            # チェックボックスの状態変数
            chk_var = ctk.StringVar(value="on" if is_selected else "off")

            # Mainクラスのメソッドを呼び出す (toggle_note_selection は Mainに残すか別途移動)
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
                fg_color=Colors.UI_BASIC,
                hover_color=Colors.adjust_brightness(Colors.UI_BASIC),
                checkmark_color=Colors.adjust_brightness(Colors.UI_BASIC, 1.8),
            )
            if not note_key:
                checkbox.configure(state="disabled")

            checkbox.pack(side="left", padx=(0, 5))

            # アイコン
            cp_key = str(row.get("commonplace_key", "")).lower()
            icon = self.key_icons.get(cp_key, "•")
            color = self.key_colors.get(cp_key, "gray")

            icon_label = ctk.CTkLabel(
                item_frame, text=icon, text_color=color, font=("", 16), width=20
            )
            icon_label.pack(side="left")

            # テキスト
            display_text = f"[{row.get('date')}] {row.get('title', 'N/A')}"
            text_label = ctk.CTkLabel(item_frame, text=display_text, anchor="w")
            text_label.pack(side="left", fill="x", expand=True)

            # --- ウィジェットリストへの保存 ---
            current_widget_index = len(self.list_item_widgets)

            self.list_item_widgets.append(
                {"frame": item_frame, "data": row, "chk_var": chk_var, "key": note_key}
            )

            # --- イベントバインド ---
            # Mainクラスのメソッド (show_details, open_pdf) を呼び出す
            def create_show_details_handler(note_row=row, idx=current_widget_index):
                def handler(event):
                    self._set_list_cursor(idx)  # ListNavigatorMixin
                    if self._details_timer:
                        self.after_cancel(self._details_timer)
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
            item_frame.bind("<Button-1>", show_details_command)
            icon_label.bind("<Button-1>", show_details_command)
            text_label.bind("<Button-1>", show_details_command)

            open_pdf_command = create_open_pdf_handler()
            item_frame.bind("<Double-Button-1>", open_pdf_command)
            icon_label.bind("<Double-Button-1>", open_pdf_command)
            text_label.bind("<Double-Button-1>", open_pdf_command)

        # リスト更新後、スクロール位置を最上部にリセットする
        self.results_list._parent_canvas.yview_moveto(0)

    def _update_pagination_ui(self):
        """ページネーションUIの更新"""
        total = getattr(self, "total_items", 0)
        current = getattr(self, "current_page", 0)
        per_page = getattr(self, "items_per_page", 50)

        max_page = max(0, (total - 1) // per_page)

        self.page_label.configure(text=f"{current + 1} / {max_page + 1}")

        if current > 0:
            self.btn_prev_page.configure(state="normal")
        else:
            self.btn_prev_page.configure(state="disabled")

        if current < max_page:
            self.btn_next_page.configure(state="normal")
        else:
            self.btn_next_page.configure(state="disabled")

    # -------------------------------------------------------------------------
    # ヘルプ表示メソッド (ボタンから呼ばれる)
    # -------------------------------------------------------------------------
    def show_search_help(self):
        """検索プレフィックスのヘルプウィンドウを表示する。"""
        # 既にウィンドウが開いている場合は、それをフォーカスする
        if hasattr(self, "help_window") and self.help_window.winfo_exists():
            self.help_window.focus()
            self.help_window.grab_set()
            return

        # カスタムクラス (SearchHelpWindow) をインスタンス化する
        self.help_window = SearchHelpWindow(self)


# ==============================================================================
# 検索ヘルプウィンドウ (Mixinファイル内にクラスとして定義)
# ==============================================================================
class SearchHelpWindow(ctk.CTkToplevel):
    """
    検索ヘルプ専用のToplevelウィンドウ。
    """

    def __init__(self, parent_app):
        super().__init__(parent_app)
        self.configure(
            fg_color=(Colors.BACKGROUND_HOLLOW, Colors.BACKGROUND_DARK_HOLLOW)
        )
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

[リスト操作 (ページネーション)]
← / → : 前のページ / 次のページへ移動
Alt + Home : 最初のページ (1ページ目) へ移動
Alt + End  : 最後のページへ移動

[リスト操作 (カーソル・スクロール)]
↑ / ↓ : リスト内をカーソル移動
Home / End : リストの先頭 / 末尾へスクロール (ページ移動はしません)
PageUp / PageDn : 画面スクロール
Ctrl + J : 表示中のノートへジャンプ (ページを自動で移動)

[リスト操作 (選択・実行)]
Space : 選択切り替え (チェックボックス ON/OFF)
Enter : 詳細を表示 (右ペイン更新)
Shift + Enter : ノートのPDFを開く
Shift + 移動 : 範囲選択 (標準)
Ctrl + Shift + 移動 : 範囲選択 (追加モード)
Ctrl + A : すべて選択 (表示ページ内)
Ctrl + D : 選択解除

[編集・アクション]
Ctrl + E : 選択中のノートを編集
Ctrl + Enter : 選択中のノートをCanvasへ送る
Alt + S  : ソート順切り替え (昇順/降順)

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
- タグを検索します (部分一致)。

`ikey: (キーワード)` (エイリアス: `cpkey:`, `indexkey:`)
- Index Key (コモンプレイスキー) を検索します (部分一致)。

`memo: (キーワード)`
- メモ欄を検索します (部分一致・「本文・メモ検索」OFFでも強制検索)。

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

2. 以降・以前 (例: `date:>=2024`, `date:<=202404`)
   年(4桁)や年月(6桁)での省略指定が可能です。
   `date:>=2024` は `20240101` 以降として扱われます。
   `date:<=2023` は `20231231` 以前として扱われます。

3. 前・後 (例: `date: >20240401`, `date: <2024`)
   年(4桁)や年月(6桁)での省略指定が可能です。
   `date:>2024` は `20250101` 以降として扱われます。
   `date:<2023` は `20221231` 以前として扱われます。

4. 期間 (例: `date:20240101-20240131`)
   2024年1月1日から1月31日までのノートを検索します。

5. 毎年 (例: `date:12-31`, `date:12-`)
   年を無視して検索します。
   `date:12-31` : 毎年12月31日のノートを検索
   `date:12-`   : 毎年12月のノートを全て検索

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

        textbox = ctk.CTkTextbox(self, wrap="word", fg_color=Colors.BACKGROUND_PANEL)
        textbox.pack(fill="both", expand=True, padx=10, pady=(10, 5))
        textbox.insert("1.0", help_text)
        textbox.configure(state="disabled")

        close_button = ctk.CTkButton(
            self,
            text="閉じる",
            command=self.destroy,
            fg_color=Colors.UI_BASIC,
            hover_color=Colors.adjust_brightness(Colors.UI_BASIC),
            text_color="black",
        )
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

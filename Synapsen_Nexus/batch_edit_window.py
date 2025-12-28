import customtkinter as ctk
from theme import SemanticColors as Colors


class BatchEditWindow(ctk.CTkToplevel):
    def __init__(
        self, parent, selection_count, all_tags, tags_in_selection, index_key_options
    ):
        super().__init__(parent)
        self.title(f"一括編集 (対象: {selection_count}件)")
        self.configure(
            fg_color=(Colors.BACKGROUND_HOLLOW, Colors.BACKGROUND_DARK_HOLLOW)
        )
        self.geometry("500x700")
        self.resizable(True, True)

        self.result = None
        self.configure(
            fg_color=(Colors.BACKGROUND_HOLLOW, Colors.BACKGROUND_DARK_HOLLOW)
        )

        def set_window(can_use_icon: bool) -> None:
            """
            アイコンを設置した後、ウィンドウを設定する。

            Args:
                can_use_icon (bool): アイコンの適応が可能か。
            """
            if can_use_icon:
                self.iconbitmap(parent.icon_path)

            # モーダル設定
            self.transient(parent)
            self.grab_set()

        # アイコン・ウィンドウ設定
        if hasattr(parent, "icon_path") and parent.icon_path:
            try:
                # アイコンの設定が可能なら、遅延して反映する。
                self.after(200, lambda: set_window(True))
            except Exception:
                # できないなら、モーダル設定のみ行う。
                set_window(False)

        # --- UI構築 ---
        self.grid_columnconfigure(0, weight=1)
        self.grid_rowconfigure(3, weight=1)  # メモ欄を伸縮可能に

        # 1. Index Key 変更
        self.lbl_ikey = ctk.CTkLabel(
            self,
            text="Index Key の変更 (選択すると上書きされます)",
            anchor="w",
            font=("", 12, "bold"),
        )
        self.lbl_ikey.grid(row=0, column=0, padx=15, pady=(15, 5), sticky="ew")

        entry_color = (
            Colors.adjust_brightness(Colors.BACKGROUND_HOLLOW, 1.1),
            Colors.adjust_brightness(Colors.BACKGROUND_DARK_HOLLOW, 1.1),
        )

        self.ikey_var = ctk.StringVar(value="変更しない")
        # 「変更しない」を先頭に追加
        options = ["変更しない"] + index_key_options
        self.ikey_combo = ctk.CTkComboBox(
            self,
            values=options,
            variable=self.ikey_var,
            state="readonly",
            width=300,
            fg_color=entry_color,
            button_color=Colors.adjust_brightness(Colors.UI_SETTING, 1.1),
            button_hover_color=Colors.UI_SETTING,
            dropdown_fg_color=(Colors.BACKGROUND_HOLLOW, Colors.BACKGROUND_DARK_HOLLOW),
            dropdown_hover_color=(
                Colors.adjust_brightness(Colors.BACKGROUND_HOLLOW, 0.85),
                Colors.adjust_brightness(Colors.BACKGROUND_DARK_HOLLOW, 0.15),
            ),
        )
        self.ikey_combo.grid(row=1, column=0, padx=15, pady=5, sticky="ew")

        # 2. タグ操作 (Tabview)
        self.lbl_tags = ctk.CTkLabel(
            self, text="タグの追加・削除", anchor="w", font=("", 12, "bold")
        )
        self.lbl_tags.grid(row=2, column=0, padx=15, pady=(15, 5), sticky="ew")

        fg_tabview = Colors.blend_colors(
            Colors.adjust_saturation(Colors.UI_TERTIARY, 0.7),
            "#FFFFFF",
            alpha=0.5,
        )
        fg_button = Colors.adjust_brightness(Colors.UI_CANCEL, 1.2)

        fg_selected = Colors.UI_BASIC
        hover_selected = Colors.adjust_brightness(fg_selected)
        hover_unselected = Colors.adjust_brightness(fg_button, 0.7)

        self.tabview = ctk.CTkTabview(
            self,
            height=200,
            fg_color=fg_tabview,
            segmented_button_fg_color=fg_button,
            segmented_button_unselected_color=fg_button,
            segmented_button_unselected_hover_color=hover_unselected,
            segmented_button_selected_color=fg_selected,
            segmented_button_selected_hover_color=hover_selected,
            text_color="black",
        )
        self.tabview.grid(row=3, column=0, padx=15, pady=5, sticky="nsew")

        # タグ選択のスクロールの色
        scroll_tag_fg = (Colors.BACKGROUND_HOLLOW, Colors.BACKGROUND_DARK_HOLLOW)
        scroll_tag_label = (
            Colors.adjust_brightness(Colors.BACKGROUND_HOLLOW),
            Colors.adjust_brightness(Colors.BACKGROUND_DARK_HOLLOW, 1.2),
        )

        # タブ: 追加
        tab_add = self.tabview.add("追加するタグ")
        self.scroll_add = ctk.CTkScrollableFrame(
            tab_add,
            label_text="既存タグから選択",
            fg_color=scroll_tag_fg,
            label_fg_color=scroll_tag_label,
        )
        self.scroll_add.pack(fill="both", expand=True, padx=5, pady=5)

        self.add_tag_vars = {}
        # 既存タグ一覧を表示（チェックボックス）
        for tag in sorted(all_tags):
            var = ctk.StringVar(value="off")
            chk = ctk.CTkCheckBox(
                self.scroll_add, text=tag, variable=var, onvalue="on", offvalue="off"
            )
            chk.pack(anchor="w", pady=2)
            self.add_tag_vars[tag] = var

        # 新規タグ入力
        self.new_tag_entry = ctk.CTkEntry(
            tab_add,
            placeholder_text="新規タグ (カンマ区切りで複数可)",
            fg_color=entry_color,
        )
        self.new_tag_entry.pack(fill="x", padx=5, pady=(10, 5))

        # タブ: 削除
        tab_remove = self.tabview.add("削除するタグ")
        self.scroll_remove = ctk.CTkScrollableFrame(
            tab_remove,
            label_text="選択中のノートに含まれるタグ",
            fg_color=scroll_tag_fg,
            label_fg_color=scroll_tag_label,
        )
        self.scroll_remove.pack(fill="both", expand=True, padx=5, pady=5)

        self.remove_tag_vars = {}
        if tags_in_selection:
            for tag in tags_in_selection:
                var = ctk.StringVar(value="off")
                chk = ctk.CTkCheckBox(
                    self.scroll_remove,
                    text=tag,
                    variable=var,
                    onvalue="on",
                    offvalue="off",
                    fg_color=Colors.LABEL_DENGER,
                    hover_color=Colors.adjust_brightness(Colors.LABEL_DENGER),
                )
                chk.pack(anchor="w", pady=2)
                self.remove_tag_vars[tag] = var
        else:
            ctk.CTkLabel(self.scroll_remove, text="共通するタグはありません").pack()

        # 3. メモ追記
        self.lbl_memo = ctk.CTkLabel(
            self,
            text="メモの追記 (各ノートの末尾に追加)",
            anchor="w",
            font=("", 12, "bold"),
        )
        self.lbl_memo.grid(row=4, column=0, padx=15, pady=(15, 5), sticky="ew")

        self.memo_textbox = ctk.CTkTextbox(
            self,
            height=80,
            fg_color=entry_color,
        )
        self.memo_textbox.grid(row=5, column=0, padx=15, pady=5, sticky="ew")

        # 4. アクションボタン
        self.btn_frame = ctk.CTkFrame(self, fg_color="transparent")
        self.btn_frame.grid(row=6, column=0, padx=15, pady=15, sticky="ew")

        self.btn_cancel = ctk.CTkButton(
            self.btn_frame,
            text="キャンセル",
            text_color=(
                Colors.adjust_brightness(Colors.BACKGROUND_PANEL, 0.2),
                Colors.adjust_brightness(Colors.BACKGROUND_DARK_PANEL, 1.8),
            ),
            fg_color="transparent",
            hover_color=(
                Colors.adjust_brightness(Colors.BACKGROUND_PANEL),
                Colors.adjust_brightness(Colors.BACKGROUND_DARK_PANEL, 1.2),
            ),
            border_width=1,
            command=self.on_cancel,
        )
        self.btn_cancel.pack(side="left", expand=True, fill="x", padx=5)

        self.btn_apply = ctk.CTkButton(
            self.btn_frame,
            text="適用する",
            fg_color=Colors.UI_EDIT,
            hover_color=Colors.adjust_brightness(Colors.UI_EDIT, 0.6),
            command=self.on_apply,
        )
        self.btn_apply.pack(side="left", expand=True, fill="x", padx=5)

    def on_apply(self):
        # Index Key
        selected_ikey = self.ikey_var.get()
        index_key_to_set = selected_ikey if selected_ikey != "変更しない" else None

        # Tags to Add
        tags_to_add = [
            tag for tag, var in self.add_tag_vars.items() if var.get() == "on"
        ]
        # 新規入力分
        new_tags_str = self.new_tag_entry.get().strip()
        if new_tags_str:
            for t in new_tags_str.split(","):
                t_clean = t.strip()
                if t_clean:
                    tags_to_add.append(t_clean)

        # Tags to Remove
        tags_to_remove = [
            tag for tag, var in self.remove_tag_vars.items() if var.get() == "on"
        ]

        # Memo
        memo_text = self.memo_textbox.get("1.0", "end-1c").strip()
        if not memo_text:
            memo_text = None

        self.result = {
            "index_key": index_key_to_set,
            "tags_to_add": tags_to_add,
            "tags_to_remove": tags_to_remove,
            "memo_text": memo_text,
            "overwrite_mode": False,  # Nexusでは安全のためデフォルト追記モード
        }
        self.destroy()

    def on_cancel(self):
        self.destroy()

    def get_input(self):
        self.master.wait_window(self)
        return self.result

import customtkinter as ctk
from tkinter import messagebox


# ==============================================================================
# データ編集ウィンドウ (書き込み可能)
# ==============================================================================
class NoteEditorWindow(ctk.CTkToplevel):
    def __init__(
            self,
            parent,
            note_data,
            commonplace_key_options,
            all_tags, save_callback
            ):
        super().__init__(parent)
        self.parent_app = parent
        self.note_data = note_data.copy()  # 元データを変更しないようコピー
        self.commonplace_key_options = commonplace_key_options
        self.all_tags = all_tags  # Nexusのpredefined_tags
        self.save_callback = save_callback  # 保存時に呼ぶ親の関数

        self.temp_tags = str(self.note_data.get("tags", "")).split(';')
        self.temp_tags = [tag for tag in self.temp_tags if tag]  # 空文字を除去

        # --- アイコン設定 ---
        self._custom_icon_path = None
        if hasattr(parent, 'icon_path') and parent.icon_path:
            self._custom_icon_path = str(parent.icon_path)
            if self._custom_icon_path:
                try:
                    super().iconbitmap(self._custom_icon_path)
                except Exception as e:
                    print(f"Initial icon set error: {e}")

        self.title(f"ノート編集: {self.note_data['title']}")
        self.geometry("500x700")
        self.transient(parent)  # 親ウィンドウの上に表示
        self.grab_set()  # モーダル化 (このウィンドウを閉じるまで親を操作不可に)

        # --- 編集可能ウィジェット ---

        # Index Key (ComboBox)
        cp_key_frame = ctk.CTkFrame(self, fg_color="transparent")
        cp_key_frame.pack(pady=10, padx=10, fill="x")
        ctk.CTkLabel(
            cp_key_frame, text="Index Key:", width=150, anchor="w"
        ).pack(side="left")
        self.cp_key_combo = ctk.CTkComboBox(
            cp_key_frame,
            values=self.commonplace_key_options
        )
        self.cp_key_combo.pack(side="left", expand=True, fill="x")
        self.cp_key_combo.set(self.note_data.get("commonplace_key", ""))

        # Key (ReadOnly)
        key_frame = ctk.CTkFrame(self, fg_color="transparent")
        key_frame.pack(pady=10, padx=10, fill="x")
        ctk.CTkLabel(
            key_frame, text="ユニークID (Key):", width=150, anchor="w"
        ).pack(side="left")
        self.key_entry = ctk.CTkEntry(key_frame)
        self.key_entry.pack(side="left", expand=True, fill="x")
        self.key_entry.insert(0, self.note_data.get("key", ""))
        self.key_entry.configure(state="readonly")

        # Memo (Textbox)
        memo_frame = ctk.CTkFrame(self, fg_color="transparent")
        memo_frame.pack(pady=10, padx=10, fill="both", expand=True)
        ctk.CTkLabel(memo_frame, text="要約・引用メモ:").pack(anchor="w")
        self.memo_textbox = ctk.CTkTextbox(memo_frame, height=150)
        self.memo_textbox.pack(fill="both", expand=True)
        self.memo_textbox.insert("1.0", self.note_data.get("memo", ""))

        # Tag (Entry + Buttons)
        tag_input_frame = ctk.CTkFrame(self, fg_color="transparent")
        tag_input_frame.pack(pady=5, padx=10, fill="x")
        ctk.CTkLabel(tag_input_frame, text="新しいタグ:").pack(side="left")
        self.tag_entry = ctk.CTkEntry(
            tag_input_frame, placeholder_text="Enterで追加"
        )
        self.tag_entry.pack(side="left", padx=5, expand=True, fill="x")
        self.tag_entry.bind("<Return>", self.add_tag_event)

        tag_button_frame = ctk.CTkFrame(self, fg_color="transparent")
        tag_button_frame.pack(pady=5, padx=10)
        ctk.CTkButton(
            tag_button_frame, text="タグを追加", command=self.add_tag_event
        ).pack(side="left", padx=5)
        ctk.CTkButton(
            tag_button_frame, text="既存タグから選択", command=self.open_tag_selector
        ).pack(side="left", padx=5)

        # Tag (Display)
        self.tags_frame = ctk.CTkScrollableFrame(self, label_text="現在のタグ")
        self.tags_frame.pack(pady=10, padx=10, fill="both", expand=True)

        # Save / Cancel Buttons
        bottom_button_frame = ctk.CTkFrame(self, fg_color="transparent")
        bottom_button_frame.pack(pady=10, side="bottom")
        ctk.CTkButton(
            bottom_button_frame, text="保存", command=self.save_and_close
        ).pack(side="left", padx=5)
        ctk.CTkButton(
            bottom_button_frame, text="キャンセル", command=self.destroy
        ).pack(side="left", padx=5)

        self.update_tags_display()

    def save_and_close(self):
        """
        保存ボタン押下時。データを辞書にまとめ、コールバックを呼ぶ。
        """
        updated_data = {
            "key": self.note_data.get("key"),  # keyは不変
            "commonplace_key": self.cp_key_combo.get().strip(),
            "memo": self.memo_textbox.get("1.0", "end-1c").strip(),
            "tags": self.temp_tags  # タグはリスト形式
        }

        try:
            # 親 (Nexus) が渡した保存関数を実行
            self.save_callback(updated_data)
            self.destroy()   # 正常に保存できたら閉じる
        except Exception as e:
            messagebox.showerror(
                "保存エラー", f"データベースの更新に失敗しました:\n{e}", parent=self)

    # --- タグ編集ヘルパー (gui_dialogs.pyから移植) ---

    def update_tags_display(self):
        for widget in self.tags_frame.winfo_children():
            widget.destroy()
        for tag in sorted(self.temp_tags):
            tag_frame = ctk.CTkFrame(self.tags_frame)
            ctk.CTkLabel(tag_frame, text=tag).pack(side="left", padx=5)
            ctk.CTkButton(
                tag_frame,
                text="x",
                width=20,
                command=lambda t=tag: self.remove_tag(t)
            ).pack(side="left", padx=5)
            tag_frame.pack(anchor="w", pady=2, fill="x")

    def add_tag_event(self, event=None):
        new_tag = self.tag_entry.get().strip()
        if new_tag:
            parts = new_tag.split('_')
            for i in range(len(parts)):
                hierarchical_tag = "_".join(parts[:i+1])
                if hierarchical_tag not in self.temp_tags:
                    self.temp_tags.append(hierarchical_tag)
        self.update_tags_display()
        self.tag_entry.delete(0, "end")

    def remove_tag(self, tag_to_remove):
        self.temp_tags.remove(tag_to_remove)
        self.update_tags_display()

    def open_tag_selector(self):
        selector = TagSelectorWindow(self, self.all_tags, self.temp_tags)
        selected_tag = selector.get_selection()
        if selected_tag:
            self.tag_entry.delete(0, "end")
            self.tag_entry.insert(0, selected_tag)
            self.add_tag_event()

    def iconbitmap(self, *args, **kwargs):
        if self._custom_icon_path:
            try:
                super().iconbitmap(self._custom_icon_path)
            except Exception:
                pass
        else:
            try:
                super().iconbitmap(*args, **kwargs)
            except Exception:
                pass


# ==============================================================================
# 既存タグ選択ウィンドウ (gui_dialogs.pyから移植)
# ==============================================================================
class TagSelectorWindow(ctk.CTkToplevel):
    def __init__(self, parent, all_tags, current_tags):
        super().__init__(parent)

        self._custom_icon_path = None
        if hasattr(
            parent, 'parent_app') and hasattr(
                parent.parent_app, 'icon_path'
                ) and parent.parent_app.icon_path:
            self._custom_icon_path = str(parent.parent_app.icon_path)
            if self._custom_icon_path:
                try:
                    super().iconbitmap(self._custom_icon_path)
                except Exception as e:
                    print(f"Initial icon set error (TagSelector): {e}")

        self.selection = None
        self.title("既存のタグを選択")
        self.geometry("300x400")
        self.transient(parent)
        self.grab_set()
        scroll_frame = ctk.CTkScrollableFrame(self)
        scroll_frame.pack(fill="both", expand=True, padx=10, pady=10)

        tags_to_show = sorted(list(set(all_tags) - set(current_tags)))

        for tag in tags_to_show:
            btn = ctk.CTkButton(
                scroll_frame,
                text=tag,
                fg_color="transparent",
                anchor="w",
                command=lambda t=tag: self.select_tag(t)
            )
            btn.pack(fill="x")

    def select_tag(self, tag):
        self.selection = tag
        self.destroy()

    def get_selection(self):
        self.master.wait_window(self)
        return self.selection

    def iconbitmap(self, *args, **kwargs):
        if self._custom_icon_path:
            try:
                super().iconbitmap(self._custom_icon_path)
            except Exception:
                pass
        else:
            try:
                super().iconbitmap(*args, **kwargs)
            except Exception:
                pass

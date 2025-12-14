import customtkinter as ctk
from tkinter import messagebox
import re

import logging

logger = logging.getLogger(__name__)


# ==============================================================================
# データ編集ウィンドウ (書き込み可能)
# ==============================================================================
class NoteEditorWindow(ctk.CTkToplevel):
    def __init__(
        self, parent, note_data, commonplace_key_options, all_tags, save_callback
    ):
        super().__init__(parent)
        self.parent_app = parent
        self.note_data = note_data.copy()  # 元データを変更しないようコピー
        self.commonplace_key_options = commonplace_key_options
        self.all_tags = all_tags  # Nexusのpredefined_tags
        self.save_callback = save_callback  # 保存時に呼ぶ親の関数

        self.temp_tags = str(self.note_data.get("tags", "")).split(";")
        self.temp_tags = [tag for tag in self.temp_tags if tag]  # 空文字を除去

        # --- アイコン設定 ---
        self._custom_icon_path = None
        if hasattr(parent, "icon_path") and parent.icon_path:
            self._custom_icon_path = str(parent.icon_path)
            if self._custom_icon_path:
                try:
                    super().iconbitmap(self._custom_icon_path)
                except Exception as e:
                    logger.error(f"Initial icon set error: {e}")

        self.title(f"ノート編集: {self.note_data['title']}")
        self.geometry("1000x700")
        self.transient(parent)
        self.grab_set()

        # --- メインコンテナ ---
        main_grid_frame = ctk.CTkFrame(self, fg_color="transparent")
        main_grid_frame.pack(fill="both", expand=True, padx=10, pady=(10, 5))

        # グリッド設定 (2カラム)
        main_grid_frame.grid_columnconfigure(0, weight=1)  # 左カラム
        main_grid_frame.grid_columnconfigure(1, weight=1)  # 右カラム
        main_grid_frame.grid_rowconfigure(0, weight=1)  # 縦方向を拡張

        # --- 左カラム (Col 0) : 主要情報 ---

        # 左カラム全体を包むフレーム
        left_frame = ctk.CTkFrame(main_grid_frame, fg_color="transparent")
        left_frame.grid(row=0, column=0, padx=(0, 10), pady=0, sticky="nsew")
        left_frame.grid_columnconfigure(0, weight=1)
        left_frame.grid_rowconfigure(5, weight=1)  # メモ欄に重み

        current_row = 0

        # 0. Index Key (ComboBox)
        cp_key_frame = ctk.CTkFrame(left_frame, fg_color="transparent")
        cp_key_frame.grid(row=current_row, column=0, pady=5, sticky="ew")
        cp_key_frame.grid_columnconfigure(1, weight=1)
        ctk.CTkLabel(cp_key_frame, text="Index Key:", width=150, anchor="w").grid(
            row=0, column=0, sticky="w"
        )
        self.cp_key_combo = ctk.CTkComboBox(
            cp_key_frame, values=self.commonplace_key_options
        )
        self.cp_key_combo.grid(row=0, column=1, sticky="ew")
        self.cp_key_combo.set(self.note_data.get("commonplace_key", ""))
        current_row += 1

        # 1. Key (ReadOnly)
        key_frame = ctk.CTkFrame(left_frame, fg_color="transparent")
        key_frame.grid(row=current_row, column=0, pady=5, sticky="ew")
        key_frame.grid_columnconfigure(1, weight=1)
        ctk.CTkLabel(key_frame, text="ユニークID (Key):", width=150, anchor="w").grid(
            row=0, column=0, sticky="w"
        )
        self.key_entry = ctk.CTkEntry(key_frame)
        self.key_entry.grid(row=0, column=1, sticky="ew")
        self.key_entry.insert(0, self.note_data.get("key", ""))
        self.key_entry.configure(state="readonly")
        current_row += 1

        # 2. Summary (概要)
        summary_outer_frame = ctk.CTkFrame(left_frame, fg_color="transparent")
        summary_outer_frame.grid(row=current_row, column=0, pady=(5, 0), sticky="ew")
        ctk.CTkLabel(summary_outer_frame, text="概要 (140文字以内):", anchor="w").pack(
            side="top", fill="x"
        )

        self.summary_entry = ctk.CTkTextbox(
            summary_outer_frame, height=75, wrap="word", activate_scrollbars=False
        )
        self.summary_entry.pack(side="top", fill="x", pady=5)

        self.summary_count_label = ctk.CTkLabel(
            summary_outer_frame, text="0 / 140", anchor="e"
        )
        self.summary_count_label.pack(side="top", fill="x")

        summary_text = self.note_data.get("summary", "")
        self.summary_entry.insert("1.0", summary_text)

        self.summary_entry.bind("<KeyRelease>", self.update_summary_count)
        self.summary_entry.bind("<KeyPress>", self.forbid_newline_input)
        self.summary_entry.bind("<<Modified>>", self.update_summary_count_modified)
        self.summary_entry.mark_set("insert", "1.0")
        self.summary_entry.edit_modified(False)
        current_row += 1

        # 3. Memo (Textbox)
        ctk.CTkLabel(left_frame, text="メモ/引用:", anchor="w").grid(
            row=current_row, column=0, pady=(10, 0), sticky="ew"
        )
        current_row += 1

        self.memo_textbox = ctk.CTkTextbox(left_frame, height=275)
        self.memo_textbox.grid(row=current_row, column=0, pady=5, sticky="nsew")
        self.memo_textbox.insert("1.0", self.note_data.get("memo", ""))
        # 行の重み設定を適用するためにダミー行を挿入
        left_frame.grid_rowconfigure(current_row, weight=1)
        current_row += 1

        # --- 右カラム (Col 1) : タグ関連 ---

        # 右カラム全体を包むフレーム
        right_frame = ctk.CTkFrame(main_grid_frame, fg_color="transparent")
        right_frame.grid(row=0, column=1, padx=(10, 0), pady=0, sticky="nsew")
        right_frame.grid_columnconfigure(0, weight=1)
        right_frame.grid_rowconfigure(2, weight=1)  # タグリストに重み

        right_row = 0

        # 4. Tag Input
        tag_input_frame = ctk.CTkFrame(right_frame, fg_color="transparent")
        tag_input_frame.grid(row=right_row, column=0, pady=5, sticky="ew")
        tag_input_frame.grid_columnconfigure(1, weight=1)

        ctk.CTkLabel(tag_input_frame, text="新しいタグ:").grid(row=0, column=0, padx=5)
        self.tag_entry = ctk.CTkEntry(tag_input_frame, placeholder_text="Enterで追加")
        self.tag_entry.grid(row=0, column=1, padx=5, sticky="ew")
        self.tag_entry.bind("<Return>", self.add_tag_event)
        right_row += 1

        # 5. Tag Buttons
        tag_button_frame = ctk.CTkFrame(right_frame, fg_color="transparent")
        tag_button_frame.grid(row=right_row, column=0, pady=5, sticky="e")
        ctk.CTkButton(
            tag_button_frame, text="タグを追加", command=self.add_tag_event
        ).pack(side="left", padx=5)
        ctk.CTkButton(
            tag_button_frame, text="既存タグから選択", command=self.open_tag_selector
        ).pack(side="left", padx=5)
        right_row += 1

        # 6. Current Tags Display
        self.tags_frame = ctk.CTkScrollableFrame(right_frame, label_text="現在のタグ")
        self.tags_frame.grid(row=right_row, column=0, pady=10, sticky="nsew")
        right_row += 1

        # --- 最下段: 保存/キャンセルボタン ---
        bottom_button_frame = ctk.CTkFrame(self, fg_color="transparent")
        bottom_button_frame.pack(pady=10, side="bottom")
        ctk.CTkButton(
            bottom_button_frame, text="保存", command=self.save_and_close
        ).pack(side="left", padx=5)
        ctk.CTkButton(
            bottom_button_frame, text="キャンセル", command=self.destroy
        ).pack(side="left", padx=5)

        self.update_tags_display()
        self.update_summary_count()

    def update_summary_count_modified(self, event=None):
        """Modifiedイベントを使用して文字数を更新する（コピペ対応）"""
        if self.summary_entry.edit_modified():
            # Textboxから内容を取得する際は "1.0" から "end-1c" (末尾の改行を除く) を指定
            content = self.summary_entry.get("1.0", "end-1c")
            length = len(content)
            max_length = 140

            display_text = f"{length} / {max_length}"
            default_color = ctk.ThemeManager.theme["CTkLabel"]["text_color"]

            if length > max_length:
                self.summary_count_label.configure(text=display_text, text_color="red")
                # 140文字を超過した場合、超過分を削除
                if length > max_length:
                    self.summary_entry.delete("1.0", "end-1c")
                    self.summary_entry.insert("1.0", content[:max_length])
                    self.summary_entry.edit_modified(False)
                    self.summary_entry.after(10, self.update_summary_count)
            else:
                self.summary_count_label.configure(
                    text=display_text, text_color=default_color
                )
            self.summary_entry.edit_modified(False)

    def update_summary_count(self, event=None):
        """KeyReleaseイベントで文字数を更新する (手打ち対応)"""
        # Textboxから内容を取得する際は "1.0" から "end-1c" (末尾の改行を除く) を指定
        content = self.summary_entry.get("1.0", "end-1c")
        length = len(content)
        max_length = 140

        display_text = f"{length} / {max_length}"
        default_color = ctk.ThemeManager.theme["CTkLabel"]["text_color"]

        if length > max_length:
            self.summary_count_label.configure(text=display_text, text_color="red")
            # 140文字を超過した場合、超過分を削除
            if event and event.keysym not in ("BackSpace", "Delete"):
                self.summary_entry.delete("1.0", "end-1c")
                self.summary_entry.insert("1.0", content[:max_length])
                self.summary_entry.after(10, self.update_summary_count)
        else:
            self.summary_count_label.configure(
                text=display_text, text_color=default_color
            )

    def forbid_newline_input(self, event):
        """Enter/Return キーが押された時に、イベントをブロックし改行入力を禁止する"""
        if event.keysym == "Return" or event.keysym == "KP_Enter":
            return "break"  # イベントを伝播させず、改行入力を禁止

    def save_and_close(self):
        """
        保存ボタン押下時。改行を削除し、データを辞書にまとめ、コールバックを呼ぶ。
        """
        # Textboxから内容を取得する際は "1.0" から "end-1c" (末尾の改行を除く) を指定
        raw_summary = self.summary_entry.get("1.0", "end-1c")

        # 1. 改行文字 (\r, \n) を全てスペースに置換
        new_summary = re.sub(r"[\r\n]+", " ", raw_summary)
        # 2. 連続するスペースを一つにまとめ、両端の空白を削除
        new_summary = re.sub(r"\s+", " ", new_summary).strip()

        # 140文字制限チェック
        if len(new_summary) > 140:
            messagebox.showerror(
                "入力エラー",
                f"概要が140文字を超えています。（現在: {len(new_summary)}文字）\n保存できません。",
                parent=self,
            )
            return

        updated_data = {
            "key": self.note_data.get("key"),
            "commonplace_key": self.cp_key_combo.get().strip(),
            "memo": self.memo_textbox.get("1.0", "end-1c").strip(),
            "tags": self.temp_tags,
            "summary": new_summary,
        }

        try:
            self.save_callback(updated_data)
            self.destroy()
        except Exception as e:
            messagebox.showerror(
                "保存エラー", f"データベースの更新に失敗しました:\n{e}", parent=self
            )

    # --- タグ編集ヘルパー (gui_dialogs.pyから移植) ---

    def update_tags_display(self):
        for widget in self.tags_frame.winfo_children():
            widget.destroy()
        for tag in sorted(self.temp_tags):
            tag_frame = ctk.CTkFrame(self.tags_frame)
            ctk.CTkLabel(tag_frame, text=tag).pack(side="left", padx=5)
            ctk.CTkButton(
                tag_frame, text="x", width=20, command=lambda t=tag: self.remove_tag(t)
            ).pack(side="left", padx=5)
            tag_frame.pack(anchor="w", pady=2, fill="x")

    def add_tag_event(self, event=None):
        new_tag = self.tag_entry.get().strip()
        if new_tag:
            parts = new_tag.split("_")
            for i in range(len(parts)):
                hierarchical_tag = "_".join(parts[: i + 1])
                if hierarchical_tag not in self.temp_tags:
                    self.temp_tags.append(hierarchical_tag)
        self.update_tags_display()
        self.tag_entry.delete(0, "end")

    def remove_tag(self, tag_to_remove):
        self.temp_tags.remove(tag_to_remove)
        self.update_tags_display()

    def open_tag_selector(self):
        def on_nexus_tag_selected(selected_tag):
            if selected_tag:
                parts = selected_tag.split("_")
                for i in range(len(parts)):
                    hierarchical_tag = "_".join(parts[: i + 1])
                    if hierarchical_tag not in self.temp_tags:
                        self.temp_tags.append(hierarchical_tag)
                self.update_tags_display()

        TagSelectorWindow(self, self.all_tags, self.temp_tags, on_nexus_tag_selected)

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
    def __init__(self, parent, all_tags, current_tags, callback=None):
        super().__init__(parent)

        self._custom_icon_path = None
        if (
            hasattr(parent, "parent_app")
            and hasattr(parent.parent_app, "icon_path")
            and parent.parent_app.icon_path
        ):
            self._custom_icon_path = str(parent.parent_app.icon_path)
            if self._custom_icon_path:
                try:
                    super().iconbitmap(self._custom_icon_path)
                except Exception as e:
                    logger.error(f"Initial icon set error (TagSelector): {e}")

        self.callback = callback
        self.title("既存のタグを選択")
        self.geometry("300x450")
        self.transient(parent)
        self.grab_set()

        scroll_frame = ctk.CTkScrollableFrame(self)
        scroll_frame.pack(fill="both", expand=True, padx=10, pady=5)

        # 常に全タグを表示するように変更（連続追加のため）
        tags_to_show = sorted(list(set(all_tags) - set(current_tags)))

        for tag in tags_to_show:
            btn = ctk.CTkButton(
                scroll_frame,
                text=tag,
                text_color=("#1F1F1F", "#1F1F1F"),
                fg_color="transparent",
                anchor="w",
                command=lambda t=tag: self.select_tag(t),
            )
            btn.configure(command=lambda t=tag, b=btn: self.select_tag(t, b))
            btn.pack(fill="x")

    def select_tag(self, tag, btn_widget):
        if self.callback:  # コールバックを実行（親画面に追加）
            self.callback(tag)
        btn_widget.destroy()  # 押されたボタンを画面から削除

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

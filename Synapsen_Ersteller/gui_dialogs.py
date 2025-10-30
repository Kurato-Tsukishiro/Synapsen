import customtkinter as ctk
from tkinter import messagebox
import datetime


# ==============================================================================
# データ編集ウィンドウ
# ==============================================================================
class DataEditorWindow(ctk.CTkToplevel):
    def __init__(self, parent, note_data, all_tags, commonplace_key_options):
        super().__init__(parent)
        self.parent = parent
        self.note_data = note_data

        self._custom_icon_path = None  # 強制設定するアイコンパス
        if hasattr(parent, 'icon_path') and parent.icon_path:
            self._custom_icon_path = str(parent.icon_path)

            # --- 初期アイコンをすぐに設定 ---
            if self._custom_icon_path:
                try:
                    # 親クラス(Toplevel)の iconbitmap を直接呼び出す
                    super().iconbitmap(self._custom_icon_path)
                except Exception as e:
                    print(f"Initial icon set error: {e}")

        self.all_tags = all_tags
        self.temp_tags = list(self.note_data.get("tags", []))
        self.commonplace_key_options = commonplace_key_options

        self.title(f"データ編集: {self.note_data['title']}")
        self.geometry("500x700")
        self.transient(parent)
        self.grab_set()

        cp_key_frame = ctk.CTkFrame(self, fg_color="transparent")
        cp_key_frame.pack(pady=10, padx=10, fill="x")
        ctk.CTkLabel(
            cp_key_frame,
            text="Index Key:",
            width=150,
            anchor="w"
        ).pack(side="left")
        self.cp_key_combo = ctk.CTkComboBox(
            cp_key_frame,
            values=self.commonplace_key_options
        )
        self.cp_key_combo.pack(side="left", expand=True, fill="x")
        self.cp_key_combo.set(self.note_data.get("commonplace_key", ""))

        key_frame = ctk.CTkFrame(self, fg_color="transparent")
        key_frame.pack(pady=10, padx=10, fill="x")
        ctk.CTkLabel(
            key_frame,
            text="ユニークID (Key):",
            width=150,
            anchor="w"
        ).pack(side="left")
        self.key_entry = ctk.CTkEntry(key_frame, placeholder_text="このノート固有のID")
        self.key_entry.pack(side="left", expand=True, fill="x")
        self.key_entry.insert(0, self.note_data.get("key", ""))
        self.key_entry.configure(state="readonly")

        memo_frame = ctk.CTkFrame(self, fg_color="transparent")
        memo_frame.pack(pady=10, padx=10, fill="both", expand=True)
        ctk.CTkLabel(memo_frame, text="要約・引用メモ:").pack(anchor="w")
        self.memo_textbox = ctk.CTkTextbox(memo_frame, height=150)
        self.memo_textbox.pack(fill="both", expand=True)
        self.memo_textbox.insert("1.0", self.note_data.get("memo", ""))

        tag_input_frame = ctk.CTkFrame(self, fg_color="transparent")
        tag_input_frame.pack(pady=5, padx=10, fill="x")
        ctk.CTkLabel(tag_input_frame, text="新しいタグ:").pack(side="left")
        self.tag_entry = ctk.CTkEntry(
            tag_input_frame,
            placeholder_text="Enterで追加"
        )
        self.tag_entry.pack(side="left", padx=5, expand=True, fill="x")
        self.tag_entry.bind("<Return>", self.add_tag_event)

        tag_button_frame = ctk.CTkFrame(self, fg_color="transparent")
        tag_button_frame.pack(pady=5, padx=10)
        ctk.CTkButton(
            tag_button_frame,
            text="タグを追加",
            command=self.add_tag_event
        ).pack(side="left", padx=5)
        ctk.CTkButton(
            tag_button_frame,
            text="既存タグから選択",
            command=self.open_tag_selector
        ).pack(side="left", padx=5)

        self.tags_frame = ctk.CTkScrollableFrame(self, label_text="現在のタグ")
        self.tags_frame.pack(pady=10, padx=10, fill="both", expand=True)

        bottom_button_frame = ctk.CTkFrame(self, fg_color="transparent")
        bottom_button_frame.pack(pady=10, side="bottom")
        ctk.CTkButton(
            bottom_button_frame,
            text="保存",
            command=self.save_and_close
        ).pack(side="left", padx=5)
        ctk.CTkButton(
            bottom_button_frame,
            text="キャンセル",
            command=self.destroy
        ).pack(side="left", padx=5)

        self.update_tags_display()

    def save_and_close(self):
        self.note_data["commonplace_key"] = self.cp_key_combo.get().strip()
        self.note_data["key"] = self.key_entry.get().strip()
        self.note_data["memo"] = self.memo_textbox.get("1.0", "end-1c").strip()
        self.note_data["tags"] = self.temp_tags
        self.parent.update_note_list()
        self.destroy()

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
        """
        iconbitmap の呼び出しをインターセプト（横取り）する。

        CustomTkinterが内部でこのメソッドを呼び出して
        アイコンをデフォルトに戻そうとしても、
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


# ==============================================================================
# 既存タグ選択ウィンドウ
# ==============================================================================
class TagSelectorWindow(ctk.CTkToplevel):
    def __init__(self, parent, all_tags, current_tags):
        super().__init__(parent)

        self._custom_icon_path = None
        
        # parent (DataEditorWindow) が 'parent' (メインアプリ) 属性を持ち、
        # かつ、その 'parent' (メインアプリ) が 'icon_path' を持っているか確認
        if hasattr(parent, 'parent') and hasattr(parent.parent, 'icon_path') and parent.parent.icon_path:
            
            # メインアプリ (parent.parent) の icon_path を直接取得
            self._custom_icon_path = str(parent.parent.icon_path)
            
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
                text_color=("#1F1F1F", "#1F1F1F"),
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
        """
        iconbitmap の呼び出しをインターセプト（横取り）する。

        CustomTkinterが内部でこのメソッドを呼び出して
        アイコンをデフォルトに戻そうとしても、
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


# ==============================================================================
# 年月入力ダイアログ
# ==============================================================================
class DateInputDialog(ctk.CTkToplevel):
    def __init__(self, parent):
        super().__init__(parent)

        self._custom_icon_path = None  # 強制設定するアイコンパス
        if hasattr(parent, 'icon_path') and parent.icon_path:
            self._custom_icon_path = str(parent.icon_path)

            # --- 初期アイコンをすぐに設定 ---
            if self._custom_icon_path:
                try:
                    # 親クラス(Toplevel)の iconbitmap を直接呼び出す
                    super().iconbitmap(self._custom_icon_path)
                except Exception as e:
                    print(f"Initial icon set error: {e}")

        self.title("年月を指定")
        self.geometry("300x200")
        self.result = None
        today = datetime.date.today()
        first_day_of_month = today.replace(day=1)
        last_month_date = first_day_of_month - datetime.timedelta(days=1)
        self.label = ctk.CTkLabel(self, text="生成するPDFの年月を入力してください:")
        self.label.pack(pady=10, padx=10)
        self.year_entry = ctk.CTkEntry(self, placeholder_text="年")
        self.year_entry.pack(pady=5)
        self.year_entry.insert(0, str(last_month_date.year))
        self.month_entry = ctk.CTkEntry(self, placeholder_text="月")
        self.month_entry.pack(pady=5)
        self.month_entry.insert(0, str(last_month_date.month))
        self.ok_button = ctk.CTkButton(self, text="OK", command=self.on_ok)
        self.ok_button.pack(pady=10)
        self.transient(parent)
        self.grab_set()

    def on_ok(self):
        try:
            year = int(self.year_entry.get())
            month = int(self.month_entry.get())
            if not (1 <= month <= 12):
                messagebox.showerror("入力エラー", "月は1から12の間で入力してください。")
                return
            self.result = (year, month)
            self.destroy()
        except ValueError:
            messagebox.showerror("入力エラー", "年と月には半角数字を入力してください。")

    def get_input(self):
        self.master.wait_window(self)
        return self.result

    def iconbitmap(self, *args, **kwargs):
        """
        iconbitmap の呼び出しをインターセプト（横取り）する。

        CustomTkinterが内部でこのメソッドを呼び出して
        アイコンをデフォルトに戻そうとしても、
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

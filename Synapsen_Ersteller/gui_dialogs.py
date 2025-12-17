import sys
from pathlib import Path
import re
import customtkinter as ctk
from tkinter import messagebox, filedialog
import datetime
import os

import logging

logger = logging.getLogger(__name__)


current_dir = Path(__file__).parent
root_dir = current_dir.parent
if str(root_dir) not in sys.path:
    sys.path.append(str(root_dir))

from theme import SemanticColors as Colors  # noqa: E402


# ==============================================================================
# データ編集ウィンドウ
# ==============================================================================
class DataEditorWindow(ctk.CTkToplevel):
    def __init__(self, parent, note_data, all_tags, commonplace_key_options):
        super().__init__(parent)
        self.parent = parent
        self.note_data = note_data

        self._custom_icon_path = None  # 強制設定するアイコンパス
        if hasattr(parent, "icon_path") and parent.icon_path:
            self._custom_icon_path = str(parent.icon_path)

            # --- 初期アイコンをすぐに設定 ---
            if self._custom_icon_path:
                try:
                    # 親クラス(Toplevel)の iconbitmap を直接呼び出す
                    super().iconbitmap(self._custom_icon_path)
                except Exception as e:
                    logger.error(f"Initial icon set error: {e}")

        self.all_tags = all_tags
        self.temp_tags = list(self.note_data.get("tags", []))
        self.commonplace_key_options = commonplace_key_options

        self.title(f"データ編集: {self.note_data['title']}")
        self.configure(
            fg_color=(Colors.BACKGROUND_HOLLOW, Colors.BACKGROUND_DARK_HOLLOW)
        )
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
            cp_key_frame,
            values=self.commonplace_key_options,
            button_color=(
                Colors.adjust_brightness(Colors.BACKGROUND_PANEL),
                Colors.adjust_brightness(Colors.BACKGROUND_DARK_PANEL, 1.2),
            ),
            button_hover_color=(
                Colors.adjust_brightness(Colors.BACKGROUND_PANEL, 0.6),
                Colors.adjust_brightness(Colors.BACKGROUND_DARK_PANEL, 1.4),
            ),
            dropdown_fg_color=(Colors.BACKGROUND_HOLLOW, Colors.BACKGROUND_DARK_HOLLOW),
            dropdown_hover_color=(
                Colors.adjust_brightness(Colors.BACKGROUND_HOLLOW, 0.85),
                Colors.adjust_brightness(Colors.BACKGROUND_DARK_HOLLOW, 0.15),
            ),
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
            summary_outer_frame,
            height=75,
            wrap="word",
            activate_scrollbars=False,
            fg_color=Colors.adjust_brightness(Colors.BACKGROUND_PANEL, 1.2),
        )
        self.summary_entry.pack(side="top", fill="x", pady=5)

        self.summary_count_label = ctk.CTkLabel(
            summary_outer_frame, text="0 / 140", anchor="e"
        )
        self.summary_count_label.pack(side="top", fill="x")

        summary_text = self.note_data.get("summary", "")
        self.summary_entry.insert("1.0", summary_text)

        # イベントバインド
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

        self.memo_textbox = ctk.CTkTextbox(
            left_frame,
            height=275,
            fg_color=Colors.adjust_brightness(Colors.BACKGROUND_PANEL, 1.2),
        )
        self.memo_textbox.grid(row=current_row, column=0, pady=5, sticky="nsew")
        self.memo_textbox.insert("1.0", self.note_data.get("memo", ""))
        # 行の重み設定を適用
        left_frame.grid_rowconfigure(current_row, weight=1)
        current_row += 1

        # --- 右カラム (Col 1) : タグ関連 ---

        # 右カラム全体を包むフレーム
        right_frame = ctk.CTkFrame(main_grid_frame, fg_color="transparent")
        right_frame.grid(row=0, column=1, padx=(10, 0), pady=0, sticky="nsew")
        right_frame.grid_columnconfigure(0, weight=1)
        right_frame.grid_rowconfigure(2, weight=1)  # タグリストに重み

        right_row = 0

        button_fg_color = Colors.UI_BASIC
        button_hover_color = Colors.adjust_brightness(Colors.UI_BASIC)

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
            tag_button_frame,
            text="タグを追加",
            fg_color=button_fg_color,
            hover_color=button_hover_color,
            text_color="black",
            command=self.add_tag_event,
        ).pack(side="left", padx=5)
        ctk.CTkButton(
            tag_button_frame,
            text="既存タグから選択",
            fg_color=button_fg_color,
            hover_color=button_hover_color,
            text_color="black",
            command=self.open_tag_selector,
        ).pack(side="left", padx=5)
        right_row += 1

        # 6. Current Tags Display
        self.tags_frame = ctk.CTkScrollableFrame(
            right_frame,
            label_text="現在のタグ",
            fg_color=Colors.BACKGROUND_PANEL,
            label_fg_color=Colors.adjust_brightness(Colors.BACKGROUND_PANEL, 0.8),
        )
        self.tags_frame.grid(row=right_row, column=0, pady=10, sticky="nsew")
        right_row += 1

        # --- 最下段: 保存/キャンセルボタン ---
        bottom_button_frame = ctk.CTkFrame(self, fg_color="transparent")
        bottom_button_frame.pack(pady=10, side="bottom")
        ctk.CTkButton(
            bottom_button_frame,
            text="保存",
            fg_color=button_fg_color,
            hover_color=button_hover_color,
            command=self.save_and_close,
            text_color="black",
        ).pack(side="left", padx=5)
        ctk.CTkButton(
            bottom_button_frame,
            text="キャンセル",
            fg_color=Colors.UI_CANCEL,
            hover_color=Colors.adjust_brightness(Colors.UI_CANCEL),
            command=self.destroy,
        ).pack(side="left", padx=5)

        self.update_tags_display()
        self.update_summary_count()  # <- 初期表示の実行 (1回だけ)

    def update_summary_count_modified(self, event=None):
        """Modifiedイベントを使用して文字数を更新する（コピペ対応）"""
        if self.summary_entry.edit_modified():
            content = self.summary_entry.get("1.0", "end-1c")
            length = len(content)
            max_length = 140

            display_text = f"{length} / {max_length}"
            default_color = ctk.ThemeManager.theme["CTkLabel"]["text_color"]

            if length > max_length:
                self.summary_count_label.configure(text=display_text, text_color="red")
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
        content = self.summary_entry.get("1.0", "end-1c")
        length = len(content)
        max_length = 140

        display_text = f"{length} / {max_length}"
        default_color = ctk.ThemeManager.theme["CTkLabel"]["text_color"]

        if length > max_length:
            self.summary_count_label.configure(text=display_text, text_color="red")
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
            return "break"

    def save_and_close(self):
        raw_summary = self.summary_entry.get("1.0", "end-1c")

        # 改行と連続スペースの削除/置換
        new_summary = re.sub(r"[\r\n]+", " ", raw_summary)
        new_summary = re.sub(r"\s+", " ", new_summary).strip()

        if len(new_summary) > 140:
            messagebox.showerror(
                "入力エラー",
                f"概要が140文字を超えています。（現在: {len(new_summary)}文字）\n保存できません。",
                parent=self,
            )
            return

        self.note_data["commonplace_key"] = self.cp_key_combo.get().strip()
        self.note_data["key"] = self.key_entry.get().strip()
        self.note_data["memo"] = self.memo_textbox.get("1.0", "end-1c").strip()
        self.note_data["tags"] = self.temp_tags
        self.note_data["summary"] = new_summary
        self.parent.update_note_list()
        self.destroy()

    def update_tags_display(self):
        for widget in self.tags_frame.winfo_children():
            widget.destroy()
        fg = Colors.UI_BASIC
        hover = Colors.adjust_brightness(Colors.UI_BASIC)
        for tag in sorted(self.temp_tags):
            tag_frame = ctk.CTkFrame(self.tags_frame)
            ctk.CTkLabel(tag_frame, text=tag).pack(side="left", padx=5)
            ctk.CTkButton(
                tag_frame,
                text="×",
                width=20,
                fg_color=fg,
                hover_color=hover,
                text_color="black",
                command=lambda t=tag: self.remove_tag(t),
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
        # タグが選択されたときに実行する関数
        def on_tag_selected(selected_tag):
            # 直接入力欄に入れて追加イベントを発火させる、または直接リストに追加する
            # ここでは直接リストに追加して表示更新するフローを採用します
            if selected_tag:
                # 階層タグの処理（既存ロジックの流用）
                parts = selected_tag.split("_")
                for i in range(len(parts)):
                    hierarchical_tag = "_".join(parts[: i + 1])
                    if hierarchical_tag not in self.temp_tags:
                        self.temp_tags.append(hierarchical_tag)

                self.update_tags_display()

        # コールバック関数を渡してウィンドウを開く
        TagSelectorWindow(self, self.all_tags, self.temp_tags, on_tag_selected)

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
    def __init__(self, parent, all_tags, current_tags, callback=None):
        super().__init__(parent)
        self.configure(
            fg_color=(Colors.BACKGROUND_HOLLOW, Colors.BACKGROUND_DARK_HOLLOW)
        )

        self._custom_icon_path = None

        # parent (DataEditorWindow) が 'parent' (メインアプリ) 属性を持ち、
        # かつ、その 'parent' (メインアプリ) が 'icon_path' を持っているか確認
        if (
            hasattr(parent, "parent")
            and hasattr(parent.parent, "icon_path")
            and parent.parent.icon_path
        ):

            # メインアプリ (parent.parent) の icon_path を直接取得
            self._custom_icon_path = str(parent.parent.icon_path)

            if self._custom_icon_path:
                try:
                    super().iconbitmap(self._custom_icon_path)
                except Exception as e:
                    logger.error(f"Initial icon set error (TagSelector): {e}")

        self.callback = callback
        self.title("既存のタグを選択")
        self.configure(
            fg_color=(Colors.BACKGROUND_HOLLOW, Colors.BACKGROUND_DARK_HOLLOW)
        )
        self.geometry("300x450")
        self.transient(parent)
        self.grab_set()

        # --- タグリスト ---
        scroll_frame = ctk.CTkScrollableFrame(
            self,
            fg_color=Colors.BACKGROUND_PANEL,
            label_fg_color=Colors.adjust_brightness(Colors.BACKGROUND_PANEL, 0.8),
        )
        scroll_frame.pack(fill="both", expand=True, padx=10, pady=5)

        tags_to_show = sorted(list(set(all_tags) - set(current_tags)))

        for tag in tags_to_show:
            btn = ctk.CTkButton(
                scroll_frame,
                text=tag,
                text_color=("#1F1F1F", "#1F1F1F"),
                fg_color="transparent",
                hover_color=Colors.adjust_brightness(Colors.BACKGROUND_PANEL, 0.8),
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

        self._custom_icon_path = None
        if hasattr(parent, "icon_path") and parent.icon_path:
            self._custom_icon_path = str(parent.icon_path)
            if self._custom_icon_path:
                try:
                    super().iconbitmap(self._custom_icon_path)
                except Exception as e:
                    logger.error(f"Initial icon set error: {e}")

        self.title("年月と表紙画像を指定")
        self.geometry("400x350")  # 高さを少し拡張
        self.result = None

        today = datetime.date.today()
        first_day_of_month = today.replace(day=1)
        last_month_date = first_day_of_month - datetime.timedelta(days=1)

        # --- 年月入力 ---
        self.label = ctk.CTkLabel(self, text="生成するPDFの年月を入力してください:")
        self.label.pack(pady=(15, 5), padx=10)

        self.date_frame = ctk.CTkFrame(self, fg_color="transparent")
        self.date_frame.pack(pady=5)

        self.year_entry = ctk.CTkEntry(self.date_frame, placeholder_text="年", width=80)
        self.year_entry.pack(side="left", padx=5)
        self.year_entry.insert(0, str(last_month_date.year))
        ctk.CTkLabel(self.date_frame, text="年").pack(side="left")

        self.month_entry = ctk.CTkEntry(
            self.date_frame, placeholder_text="月", width=60
        )
        self.month_entry.pack(side="left", padx=5)
        self.month_entry.insert(0, str(last_month_date.month))
        ctk.CTkLabel(self.date_frame, text="月").pack(side="left")

        # --- 表紙画像選択 (追加) ---
        ctk.CTkLabel(self, text="表紙画像 (任意):", text_color="gray").pack(
            pady=(20, 5), padx=10
        )

        self.img_frame = ctk.CTkFrame(self, fg_color="transparent")
        self.img_frame.pack(pady=0, padx=20, fill="x")

        self.img_path_entry = ctk.CTkEntry(
            self.img_frame, placeholder_text="画像ファイルを選択..."
        )
        self.img_path_entry.pack(side="left", fill="x", expand=True, padx=(0, 5))

        ctk.CTkButton(
            self.img_frame, text="参照", width=60, command=self.browse_image
        ).pack(side="right")

        # --- OKボタン ---
        self.ok_button = ctk.CTkButton(self, text="OK", command=self.on_ok, width=100)
        self.ok_button.pack(pady=20)

        self.transient(parent)
        self.grab_set()

    def browse_image(self):
        file_types = [("Images", "*.png;*.jpg;*.jpeg"), ("All Files", "*.*")]
        path = filedialog.askopenfilename(title="表紙画像を選択", filetypes=file_types)
        if path:
            self.img_path_entry.delete(0, "end")
            self.img_path_entry.insert(0, path)

    def on_ok(self):
        try:
            year = int(self.year_entry.get())
            month = int(self.month_entry.get())
            if not (1 <= month <= 12):
                messagebox.showerror(
                    "入力エラー", "月は1から12の間で入力してください。"
                )
                return

            # 画像パスの取得 (空欄ならNone)
            img_path = self.img_path_entry.get().strip()
            if img_path == "":
                img_path = None
            elif not os.path.exists(img_path):
                messagebox.showwarning(
                    "警告",
                    "指定された画像ファイルが見つかりません。\n画像なしで続行します。",
                )
                img_path = None

            # 結果に画像パスを含める (year, month, img_path)
            self.result = (year, month, img_path)
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


# ==============================================================================
# 既存タグ選択ウィンドウ (BatchEditWindow用)
# ==============================================================================
class BatchTagSelectorWindow(ctk.CTkToplevel):
    """
    BatchEditWindow が「追加するタグ」を選択するために使用するウィンドウ。
    DataEditorWindow が使う TagSelectorWindow とほぼ同じだが、
    アイコンの参照元が異なるため別クラスとして定義する。
    """

    def __init__(self, parent, all_tags, current_tags, callback=None):
        super().__init__(parent)

        self._custom_icon_path = None

        # parent (BatchEditWindow) が 'parent' (メインアプリ) 属性を持ち、
        # かつ、その 'parent' (メインアプリ) が 'icon_path' を持っているか確認
        if (
            hasattr(parent, "parent")
            and hasattr(parent.parent, "icon_path")
            and parent.parent.icon_path
        ):

            # メインアプリ (parent.parent) の icon_path を直接取得
            self._custom_icon_path = str(parent.parent.icon_path)

            if self._custom_icon_path:
                try:
                    super().iconbitmap(self._custom_icon_path)
                except Exception as e:
                    logger.error("Initial icon set error (BatchTagSelector): " f"{e}")

        self.callback = callback
        self.title("既存のタグを選択")
        self.configure(
            fg_color=(Colors.BACKGROUND_HOLLOW, Colors.BACKGROUND_DARK_HOLLOW)
        )
        self.geometry("300x450")
        self.transient(parent)
        self.grab_set()

        scroll_frame = ctk.CTkScrollableFrame(
            self,
            fg_color=Colors.BACKGROUND_PANEL,
            label_fg_color=Colors.adjust_brightness(Colors.BACKGROUND_PANEL, 0.8),
        )
        scroll_frame.pack(fill="both", expand=True, padx=10, pady=5)

        tags_to_show = sorted(list(set(all_tags) - set(current_tags)))
        for tag in tags_to_show:
            btn = ctk.CTkButton(
                scroll_frame,
                text=tag,
                text_color=("#1F1F1F", "#1F1F1F"),
                fg_color="transparent",
                hover_color=Colors.adjust_brightness(Colors.BACKGROUND_PANEL, 0.8),
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


# ==============================================================================
# 一括編集ウィンドウ
# ==============================================================================
class BatchEditWindow(ctk.CTkToplevel):
    def __init__(
        self,
        parent,
        selected_count,
        all_tags,
        tags_in_selection,
        commonplace_key_options,
    ):
        super().__init__(parent)
        self.parent = parent
        self.all_tags = all_tags
        self.tags_in_selection = tags_in_selection  # 選択範囲内のタグ
        self.commonplace_key_options = commonplace_key_options
        self.result = None

        # 一時的なタグリスト
        self.tags_to_add = []
        self.tags_to_remove = []  # 削除対象のタグリスト

        self._custom_icon_path = None
        if hasattr(parent, "icon_path") and parent.icon_path:
            self._custom_icon_path = str(parent.icon_path)
            if self._custom_icon_path:
                try:
                    super().iconbitmap(self._custom_icon_path)
                except Exception as e:
                    logger.error(f"Initial icon set error: {e}")

        self.title(f"一括編集 ({selected_count} 件)")
        self.geometry("600x825")
        self.configure(
            fg_color=(Colors.BACKGROUND_HOLLOW, Colors.BACKGROUND_DARK_HOLLOW)
        )
        self.transient(parent)
        self.grab_set()

        # --- 1. Index Key ---
        cp_key_frame = ctk.CTkFrame(
            self, fg_color=(Colors.BACKGROUND_PANEL, Colors.BACKGROUND_DARK_PANEL)
        )
        cp_key_frame.pack(pady=10, padx=10, fill="x")
        ctk.CTkLabel(
            cp_key_frame,
            text="Index Key を設定:",
            width=150,
            anchor="w",
        ).pack(side="left")

        # [変更しない] オプションを追加
        cp_key_values = ["[ 変更しない ]"] + self.commonplace_key_options
        self.cp_key_combo = ctk.CTkComboBox(
            cp_key_frame,
            values=cp_key_values,
            button_color=(
                Colors.adjust_brightness(Colors.BACKGROUND_PANEL),
                Colors.adjust_brightness(Colors.BACKGROUND_DARK_PANEL, 1.2),
            ),
            button_hover_color=(
                Colors.adjust_brightness(Colors.BACKGROUND_PANEL, 0.6),
                Colors.adjust_brightness(Colors.BACKGROUND_DARK_PANEL, 1.4),
            ),
        )
        self.cp_key_combo.pack(side="left", expand=True, fill="x")
        self.cp_key_combo.set("[ 変更しない ]")  # デフォルト

        # --- 2. メモ追記 ---
        memo_frame = ctk.CTkFrame(
            self, fg_color=(Colors.BACKGROUND_PANEL, Colors.BACKGROUND_DARK_PANEL)
        )
        memo_frame.pack(pady=5, padx=10, fill="x")

        # ヘッダー (ラベルとスイッチ)
        memo_header_frame = ctk.CTkFrame(memo_frame, fg_color="transparent")
        memo_header_frame.pack(side="top", fill="x", padx=5, pady=(5, 0))

        ctk.CTkLabel(memo_header_frame, text="メモ編集 (引用など):", anchor="w").pack(
            side="left"
        )

        # 上書き切り替えスイッチ (デフォルトOFF=追記)
        self.overwrite_switch = ctk.CTkSwitch(
            memo_header_frame,
            text="上書きモード",
            button_color=Colors.adjust_brightness(Colors.UI_CANCEL),
            button_hover_color=Colors.adjust_brightness(Colors.UI_CANCEL, 0.6),
            fg_color=Colors.UI_CANCEL,
            progress_color=Colors.UI_BASIC,
        )
        self.overwrite_switch.pack(side="right")

        # テキストボックス
        self.memo_input_box = ctk.CTkTextbox(
            memo_frame,
            height=80,
            fg_color=(
                Colors.adjust_brightness(Colors.BACKGROUND_PANEL, 1.2),
                Colors.adjust_brightness(Colors.BACKGROUND_DARK_PANEL),
            ),
        )
        self.memo_input_box.pack(side="bottom", fill="x", padx=5, pady=5)

        # --- 3. 追加するタグ ---
        add_tag_frame = ctk.CTkFrame(
            self, fg_color=(Colors.BACKGROUND_PANEL, Colors.BACKGROUND_DARK_PANEL)
        )
        add_tag_frame.pack(pady=10, padx=10, fill="both", expand=True)

        ctk.CTkLabel(add_tag_frame, text="追加するタグ:").pack(anchor="w")

        add_tag_input_frame = ctk.CTkFrame(add_tag_frame, fg_color="transparent")
        add_tag_input_frame.pack(pady=5, fill="x")
        self.add_tag_entry = ctk.CTkEntry(
            add_tag_input_frame,
            placeholder_text="Enterで追加",
            fg_color=(
                Colors.adjust_brightness(Colors.BACKGROUND_PANEL, 1.2),
                Colors.adjust_brightness(Colors.BACKGROUND_DARK_PANEL),
            ),
        )
        self.add_tag_entry.pack(side="left", padx=(0, 5), expand=True, fill="x")
        self.add_tag_entry.bind("<Return>", self.add_tag_to_add_list)
        ctk.CTkButton(
            add_tag_input_frame,
            text="既存タグから選択",
            command=self.open_tag_selector_for_add,
            fg_color=Colors.UI_BASIC,
            hover_color=Colors.adjust_brightness(Colors.UI_BASIC),
            text_color="black",
        ).pack(side="left")

        self.add_tags_display_frame = ctk.CTkScrollableFrame(
            add_tag_frame,
            fg_color=(
                Colors.adjust_brightness(Colors.BACKGROUND_HOLLOW, 0.9),
                Colors.adjust_brightness(Colors.BACKGROUND_DARK_HOLLOW, 1.1),
            ),
        )
        self.add_tags_display_frame.pack(fill="both", expand=True)

        # --- 4. 削除するタグ (選択式リスト) ---
        remove_tag_frame = ctk.CTkFrame(
            self, fg_color=(Colors.BACKGROUND_PANEL, Colors.BACKGROUND_DARK_PANEL)
        )
        remove_tag_frame.pack(pady=10, padx=10, fill="both", expand=True)

        header_frame = ctk.CTkFrame(remove_tag_frame, fg_color="transparent")
        header_frame.pack(fill="x", pady=5)
        ctk.CTkLabel(
            header_frame,
            text="含まれているタグ (×で削除指定):",
            text_color=Colors.LABEL_DENGER,
        ).pack(side="left", padx=5)

        # 削除タグ一覧を表示するスクロールフレーム
        self.remove_tags_scroll = ctk.CTkScrollableFrame(
            remove_tag_frame,
            height=200,
            fg_color=(
                Colors.adjust_brightness(Colors.BACKGROUND_HOLLOW, 0.9),
                Colors.adjust_brightness(Colors.BACKGROUND_DARK_HOLLOW, 1.1),
            ),
        )
        self.remove_tags_scroll.pack(fill="both", expand=True, padx=5, pady=5)

        # 選択範囲のタグ一覧を描画
        self.tag_widgets = {}  # {tag_name: (label, button)}
        self.populate_remove_tag_list()

        # --- 4. 適用 / キャンセルボタン ---
        bottom_button_frame = ctk.CTkFrame(self, fg_color="transparent")
        bottom_button_frame.pack(pady=10, side="bottom")
        ctk.CTkButton(
            bottom_button_frame,
            text="適用",
            command=self.apply_changes,
            fg_color=Colors.UI_BASIC,
            hover_color=Colors.adjust_brightness(Colors.UI_BASIC),
            text_color="black",
        ).pack(side="left", padx=5)
        ctk.CTkButton(
            bottom_button_frame,
            text="キャンセル",
            command=self.destroy,
            fg_color=Colors.UI_CANCEL,
            hover_color=Colors.adjust_brightness(Colors.UI_CANCEL),
        ).pack(side="left", padx=5)

    def get_input(self):
        self.master.wait_window(self)
        return self.result

    def apply_changes(self):
        index_key_to_set = self.cp_key_combo.get()
        if index_key_to_set == "[ 変更しない ]":
            index_key_to_set = None

        # 入力内容とモードを取得
        text_val = self.memo_input_box.get("1.0", "end-1c").strip()
        memo_text = text_val if text_val else None

        # スイッチの状態を取得 (1=ON, 0=OFF)
        overwrite_mode = bool(self.overwrite_switch.get())

        self.result = {
            "index_key": index_key_to_set,
            "tags_to_add": self.tags_to_add,
            "tags_to_remove": self.tags_to_remove,
            "memo_text": memo_text,  # テキスト
            "overwrite_mode": overwrite_mode,  # モード
        }
        self.destroy()

    # --- 「追加するタグ」リストの管理 ---
    def add_tag_to_add_list(self, event=None):
        new_tag = self.add_tag_entry.get().strip()
        if new_tag:
            parts = new_tag.split("_")
            for i in range(len(parts)):
                hierarchical_tag = "_".join(parts[: i + 1])
                if hierarchical_tag not in self.tags_to_add:
                    self.tags_to_add.append(hierarchical_tag)
        self.update_add_tags_display()
        self.add_tag_entry.delete(0, "end")

    def remove_tag_from_add_list(self, tag_to_remove):
        self.tags_to_add.remove(tag_to_remove)
        self.update_add_tags_display()

    def update_add_tags_display(self):
        for widget in self.add_tags_display_frame.winfo_children():
            widget.destroy()
        for tag in sorted(self.tags_to_add):
            tag_frame = ctk.CTkFrame(self.add_tags_display_frame)
            ctk.CTkLabel(tag_frame, text=tag).pack(side="left", padx=5)
            ctk.CTkButton(
                tag_frame,
                text="x",
                width=20,
                fg_color=Colors.UI_BASIC,
                hover_color=Colors.adjust_brightness(Colors.UI_BASIC),
                command=lambda t=tag: self.remove_tag_from_add_list(t),
            ).pack(side="left", padx=5)
            tag_frame.pack(anchor="w", pady=2, fill="x")

    def open_tag_selector_for_add(self):
        def on_batch_tag_selected(selected_tag):
            if selected_tag:
                parts = selected_tag.split("_")
                for i in range(len(parts)):
                    hierarchical_tag = "_".join(parts[: i + 1])
                    if hierarchical_tag not in self.tags_to_add:
                        self.tags_to_add.append(hierarchical_tag)

                self.update_add_tags_display()

        # BatchTagSelectorWindowをコールバック付きで呼び出し
        BatchTagSelectorWindow(
            self, self.all_tags, self.tags_to_add, on_batch_tag_selected
        )

    # --- 「削除するタグ」リストの管理 ---
    def populate_remove_tag_list(self):
        """選択範囲に含まれるタグを一覧表示し、削除ボタンを配置する"""
        if not self.tags_in_selection:
            ctk.CTkLabel(
                self.remove_tags_scroll,
                text="（選択されたノートに共通するタグはありません）",
                text_color="gray",
            ).pack(pady=10)
            return

        for tag in self.tags_in_selection:
            row = ctk.CTkFrame(self.remove_tags_scroll, fg_color="transparent")
            row.pack(fill="x", pady=2)

            # タグ名ラベル
            lbl = ctk.CTkLabel(row, text=tag, anchor="w")
            lbl.pack(side="left", padx=5, fill="x", expand=True)

            # 削除/戻すボタン
            # lambdaで変数をキャプチャする際、デフォルト引数を使うことで現在の値を固定する
            btn = ctk.CTkButton(
                row,
                text="×",
                width=40,
                fg_color=Colors.LABEL_DENGER,
                hover_color=Colors.adjust_brightness(Colors.LABEL_DENGER),
                command=lambda t=tag: self.toggle_remove_tag(t),
            )
            btn.pack(side="right", padx=5)

            # ウィジェットへの参照を保持 (後で色やテキストを変えるため)
            self.tag_widgets[tag] = (lbl, btn)

    def toggle_remove_tag(self, tag):
        """タグの削除状態を切り替える (Keep <-> Remove)"""
        lbl, btn = self.tag_widgets[tag]

        if tag in self.tags_to_remove:
            # 既に削除対象 -> 元に戻す
            self.tags_to_remove.remove(tag)

            # UIを通常状態に戻す
            lbl.configure(
                text_color=ctk.ThemeManager.theme["CTkLabel"]["text_color"]
            )  # デフォルト色
            btn.configure(
                text="×",
                fg_color=Colors.LABEL_DENGER,
                hover_color=Colors.adjust_brightness(Colors.LABEL_DENGER),
            )

        else:
            # 削除対象に追加
            self.tags_to_remove.append(tag)

            # UIを削除待機状態にする
            lbl.configure(text_color="gray")  # グレーアウト
            btn.configure(text="戻す", fg_color="gray", hover_color="#555")

    def remove_tag_from_remove_list(self, tag_to_remove):
        self.tags_to_remove.remove(tag_to_remove)
        self.update_remove_tags_display()

    def update_remove_tags_display(self):
        for widget in self.remove_tags_display_frame.winfo_children():
            widget.destroy()
        fg = Colors.UI_BASIC
        hover = Colors.adjust_brightness(Colors.UI_BASIC)
        for tag in sorted(self.tags_to_remove):
            tag_frame = ctk.CTkFrame(self.remove_tags_display_frame)
            ctk.CTkLabel(tag_frame, text=tag).pack(side="left", padx=5)
            ctk.CTkButton(
                tag_frame,
                text="×",
                width=20,
                fg_color=fg,
                hover_color=hover,
                text_color="black",
                command=lambda t=tag: self.remove_tag_from_remove_list(t),
            ).pack(side="left", padx=5)
            tag_frame.pack(anchor="w", pady=2, fill="x")

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

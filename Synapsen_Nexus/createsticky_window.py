import logging
import sys
from pathlib import Path

import customtkinter as ctk


# --- プロジェクト内モジュールのパス設定 ---
current_dir = Path(__file__).parent
root_dir = current_dir.parent
normalisierer_dir = root_dir / "Synapsen_Normalisierer"

if str(root_dir) not in sys.path:
    sys.path.append(str(root_dir))
if str(normalisierer_dir) not in sys.path:
    sys.path.append(str(normalisierer_dir))

# --- プロジェクト内モジュールのインポート ---
from logging_setup import setup_logging  # noqa: E402
from theme import SemanticColors as Colors  # noqa: E402


# --- ロガー設定 ---
logger = logging.getLogger("Nexus")
if not logger.handlers:
    setup_logging("Synapsen_Nexus")


class CreateStickyDialog(ctk.CTkToplevel):
    """付箋ノート化するファイルと設定を選択するダイアログ"""

    def __init__(self, parent, key_options, sticky_colors):
        super().__init__(parent)
        self.parent_app = parent

        # --- アイコン設定 ---
        self._custom_icon_path = None
        if hasattr(parent, "icon_path") and parent.icon_path:
            self._custom_icon_path = str(parent.icon_path)
            if self._custom_icon_path:
                try:
                    super().iconbitmap(self._custom_icon_path)
                except Exception as e:
                    logger.error(f"Initial icon set error: {e}")

        self.title("付箋ノート作成")
        self.geometry("500x360")
        self.configure(
            fg_color=(Colors.BACKGROUND_HOLLOW, Colors.BACKGROUND_DARK_HOLLOW)
        )
        self.transient(parent)
        self.grab_set()

        self.result = None
        self.file_path = None

        self.grid_columnconfigure(0, weight=1)

        # ファイル選択 UI
        ctk.CTkLabel(self, text="付箋にするファイル (MD, TXT等):", anchor="w").grid(
            row=0, column=0, padx=20, pady=(20, 5), sticky="ew"
        )
        self.file_btn = ctk.CTkButton(
            self,
            text="ファイルを選択",
            command=self.select_file,
            fg_color=Colors.UI_SECONDARY,
            hover_color=Colors.adjust_brightness(Colors.UI_SECONDARY),
        )
        self.file_btn.grid(row=1, column=0, padx=20, pady=(0, 5), sticky="ew")

        self.file_label = ctk.CTkLabel(self, text="未選択", text_color="gray")
        self.file_label.grid(row=2, column=0, padx=20, pady=(0, 15), sticky="ew")

        # Index Key UI
        ctk.CTkLabel(self, text="Index Key:", anchor="w").grid(
            row=3, column=0, padx=20, pady=(0, 5), sticky="ew"
        )
        self.key_combo = ctk.CTkComboBox(
            self,
            values=key_options,
            fg_color=Colors.adjust_brightness(Colors.BACKGROUND_HOLLOW, 1.2),
            button_color=Colors.adjust_brightness(Colors.UI_SETTING, 1.1),
            button_hover_color=Colors.UI_SETTING,
            dropdown_fg_color=(Colors.BACKGROUND_HOLLOW, Colors.BACKGROUND_DARK_HOLLOW),
            dropdown_hover_color=(
                Colors.adjust_brightness(Colors.BACKGROUND_HOLLOW, 0.85),
                Colors.adjust_brightness(Colors.BACKGROUND_DARK_HOLLOW, 0.15),
            ),
        )
        self.key_combo.grid(row=4, column=0, padx=20, pady=(0, 15), sticky="ew")
        if key_options:
            self.key_combo.set(key_options[0])

        # カラー選択 UI
        ctk.CTkLabel(self, text="背景色:", anchor="w").grid(
            row=5, column=0, padx=20, pady=(0, 5), sticky="ew"
        )
        self.color_var = ctk.StringVar(
            value=sticky_colors[0][1] if sticky_colors else "#FFFFA5"
        )
        color_frame = ctk.CTkFrame(self, fg_color="transparent")
        color_frame.grid(row=6, column=0, padx=20, pady=(0, 15), sticky="ew")

        # 実行環境のテーマに応じたテキストカラー設定
        text_color = "gray90" if ctk.get_appearance_mode() == "Dark" else "black"

        for i, (name, code) in enumerate(sticky_colors):
            btn = ctk.CTkRadioButton(
                color_frame,
                text=name,
                value=code,
                variable=self.color_var,
                fg_color=code,
                text_color=text_color,
            )
            r, c = divmod(i, 3)
            btn.grid(row=r, column=c, padx=5, pady=5, sticky="w")

        # ボタン
        btn_frame = ctk.CTkFrame(self, fg_color="transparent")
        btn_frame.grid(row=7, column=0, padx=20, pady=10, sticky="e")
        ctk.CTkButton(
            btn_frame,
            text="キャンセル",
            command=self.destroy,
            width=80,
            fg_color=Colors.UI_CANCEL,
            hover_color=Colors.adjust_brightness(Colors.UI_CANCEL),
        ).pack(side="left", padx=5)

        ctk.CTkButton(
            btn_frame,
            text="作成",
            command=self.on_ok,
            width=80,
            fg_color=Colors.UI_BASIC,
            hover_color=Colors.adjust_brightness(Colors.UI_BASIC),
            text_color="black",
        ).pack(side="left", padx=5)

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

    def select_file(self):
        from tkinter import filedialog

        path = filedialog.askopenfilename(
            filetypes=[("Text/Markdown", "*.txt;*.md"), ("All files", "*.*")]
        )
        if path:
            self.file_path = path
            from pathlib import Path

            self.file_label.configure(text=Path(path).name)

    def on_ok(self):
        if not self.file_path:
            from tkinter import messagebox

            messagebox.showerror("エラー", "ファイルを選択してください。")
            return
        self.result = {
            "file_path": self.file_path,
            "index_key": self.key_combo.get(),
            "bg_color": self.color_var.get(),
        }
        self.destroy()

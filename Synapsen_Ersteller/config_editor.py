import sys
import customtkinter as ctk
import configparser
from tkinter import filedialog, messagebox
from pathlib import Path
import os


current_dir = Path(__file__).parent
root_dir = current_dir.parent
if str(root_dir) not in sys.path:
    sys.path.append(str(root_dir))

from theme import SemanticColors as Colors  # noqa: E402


class ConfigEditorWindow(ctk.CTkToplevel):
    def __init__(self, parent, config_path):
        super().__init__(parent)
        self.parent = parent
        self.config_path = Path(config_path)

        # アイコン設定 (親から継承)
        if hasattr(parent, "icon_path") and parent.icon_path:
            try:
                self.iconbitmap(default=str(parent.icon_path))
            except Exception:
                pass

        self.title("Synapsen 設定")
        self.configure(
            fg_color=(Colors.BACKGROUND_HOLLOW, Colors.BACKGROUND_DARK_HOLLOW)
        )
        self.geometry("650x800")

        # モーダル設定
        self.transient(parent)
        self.grab_set()

        # ConfigParserの初期化
        self.config = configparser.ConfigParser(interpolation=None)
        self.config.optionxform = str

        if self.config_path.exists():
            self.config.read(self.config_path, encoding="utf-8")
        else:
            messagebox.showerror(
                "エラー", f"設定ファイルが見つかりません: {self.config_path}"
            )
            self.destroy()
            return

        self._create_widgets()

    def _create_widgets(self):
        # スクロール可能なメインフレーム
        self.scroll_frame = ctk.CTkScrollableFrame(
            self,
            fg_color=Colors.BACKGROUND_PANEL,
            label_fg_color=Colors.adjust_brightness(Colors.BACKGROUND_PANEL, 0.8),
        )
        self.scroll_frame.pack(fill="both", expand=True, padx=10, pady=10)

        # --- [Paths] セクション ---
        self._add_section_label("パス設定 (Paths)")

        self.entry_db_path = self._add_path_entry(
            "Database Path:", "Paths", "database_path", is_file=True
        )
        self.entry_pdf_root = self._add_path_entry(
            "PDF Root Folder:", "Paths", "pdf_root_folder", is_file=False
        )
        self.entry_pdf_archive = self._add_path_entry(
            "PDF Archive Folder:", "Paths", "pdf_archive_folder", is_file=False
        )
        self.entry_font = self._add_path_entry(
            "基本フォント (正規化・共通):", "Paths", "font_path", is_file=True
        )

        # --- [Automation] セクション ---
        self._add_section_label("自動処理 (Automation)")

        self.var_ocr = self._add_checkbox(
            "Tesseract OCR を有効化", "Automation", "enable_tesseract_ocr"
        )
        self.var_flatten = self._add_checkbox(
            "インク注釈をフラット化", "Automation", "flatten_ink_annotations"
        )
        self.var_auto_db = self._add_checkbox(
            "DBへの自動追記 (Ersteller => Nexus)",
            "Automation",
            "auto_append_to_default_db",
        )
        self.var_export_csv = self._add_checkbox(
            "統合時のリストCSVの出力", "Automation", "create_individual_csv"
        )
        self.var_onthisday = self._add_checkbox(
            "Nexus起動時に「過去の今日」のノートを通知するか",
            "Automation",
            "enable_on_this_day",
        )

        # --- [Watchdog] セクション (New) ---
        self._add_section_label("監視設定 (Watchdog)")

        self.entry_watch_dir = self._add_path_entry(
            "監視フォルダ (Watch Input):", "Watchdog", "watch_dir", is_file=False
        )
        self.entry_watch_output = self._add_path_entry(
            "出力フォルダ (Watch Output):", "Watchdog", "output_dir", is_file=False
        )
        self.entry_watch_size = self._add_text_entry(
            "正規化サイズ (A4/A5):", "Watchdog", "target_size"
        )

        # --- [ReportLab] セクション ---
        self._add_section_label("PDF生成設定 (ReportLab)")

        self.entry_paper_size = self._add_text_entry(
            "用紙サイズ (A4/A5):", "ReportLab", "paper_size"
        )
        self.entry_rl_font = self._add_path_entry(
            "統合PDF用フォント (空欄時は基本フォント):",
            "ReportLab",
            "font",
            is_file=True,
        )
        self.entry_author = self._add_text_entry("著者名:", "ReportLab", "author")
        self.entry_title_prefix = self._add_text_entry(
            "タイトル接頭辞:", "ReportLab", "title_prefix"
        )
        self.entry_max_pages = self._add_text_entry(
            "最大ページ数 (0=無制限):", "ReportLab", "max_pages"
        )

        # --- 保存ボタン ---
        btn_frame = ctk.CTkFrame(self, fg_color="transparent")
        btn_frame.pack(fill="x", padx=20, pady=10)

        ctk.CTkButton(
            btn_frame,
            text="保存して閉じる",
            command=self._save_config,
            fg_color=Colors.UI_EDIT,
            hover_color=Colors.adjust_brightness(Colors.UI_EDIT, 0.6),
        ).pack(side="right", padx=5)
        ctk.CTkButton(
            btn_frame,
            text="キャンセル",
            command=self.destroy,
            fg_color=Colors.UI_CANCEL,
            hover_color=Colors.adjust_brightness(Colors.UI_CANCEL),
        ).pack(side="right", padx=5)

    def _add_section_label(self, text):
        ctk.CTkLabel(
            self.scroll_frame, text=text, font=("", 16, "bold"), anchor="w"
        ).pack(fill="x", pady=(15, 5))

    def _add_path_entry(self, label_text, section, option, is_file=True):
        """パス選択用の行を作成"""
        frame = ctk.CTkFrame(self.scroll_frame, fg_color="transparent")
        frame.pack(fill="x", pady=2)

        ctk.CTkLabel(frame, text=label_text, width=200, anchor="w").pack(side="left")

        current_val = self._get_config_val(section, option)
        entry = ctk.CTkEntry(
            frame, fg_color=Colors.adjust_brightness(Colors.BACKGROUND_HOLLOW, 1.2)
        )
        entry.insert(0, current_val)
        entry.pack(side="left", fill="x", expand=True, padx=5)

        def browse():
            # --- 初期ディレクトリ(initialdir) の決定ロジック ---
            initial_dir = None
            current_entry_val = entry.get().strip()

            # 1. 現在入力されているパスが存在する場合、そのフォルダを開く
            if current_entry_val:
                path_obj = Path(current_entry_val)
                if path_obj.is_file():
                    initial_dir = path_obj.parent
                elif path_obj.is_dir():
                    initial_dir = path_obj
                elif path_obj.parent.exists():
                    initial_dir = path_obj.parent

            # 2. 入力が空、かつフォントパスの設定項目である場合
            if not initial_dir and (option == "font_path" or option == "font"):
                # 候補リスト (優先順を変更: ユーザーローカル -> システム)
                candidates = [
                    os.path.expandvars(r"%LOCALAPPDATA%\Microsoft\Windows\Fonts"),
                    r"C:\Windows\Fonts",
                ]
                for c in candidates:
                    if os.path.exists(c):
                        initial_dir = c
                        break

            # --- ダイアログを開く ---
            if is_file:
                # フォントファイルを見つけやすくするためのフィルタ
                file_types = [("All Files", "*.*")]
                if option == "font_path" or option == "font":
                    file_types = [("Font Files", "*.ttf;*.ttc;"), ("All Files", "*.*")]

                path = filedialog.askopenfilename(
                    initialdir=initial_dir, filetypes=file_types
                )
            else:
                path = filedialog.askdirectory(initialdir=initial_dir)

            if path:
                # パスをOS標準の区切り文字に修正して入力
                path = str(Path(path).absolute())
                entry.delete(0, "end")
                entry.insert(0, path)

        ctk.CTkButton(
            frame,
            text="参照",
            width=50,
            command=browse,
            fg_color=Colors.UI_BASIC,
            hover_color=Colors.adjust_brightness(Colors.UI_BASIC),
            text_color="black",
        ).pack(side="left")
        return entry

    def _add_text_entry(self, label_text, section, option):
        """通常のテキスト入力"""
        frame = ctk.CTkFrame(self.scroll_frame, fg_color="transparent")
        frame.pack(fill="x", pady=2)
        ctk.CTkLabel(frame, text=label_text, width=200, anchor="w").pack(side="left")

        current_val = self._get_config_val(section, option)
        entry = ctk.CTkEntry(
            frame, fg_color=Colors.adjust_brightness(Colors.BACKGROUND_HOLLOW, 1.2)
        )
        entry.insert(0, current_val)
        entry.pack(side="left", fill="x", expand=True)
        return entry

    def _add_checkbox(self, label_text, section, option):
        """Boolean用チェックボックス"""
        frame = ctk.CTkFrame(self.scroll_frame, fg_color="transparent")
        frame.pack(fill="x", pady=2)

        val_str = self._get_config_val(section, option).lower()
        is_checked = val_str == "true"

        var = ctk.BooleanVar(value=is_checked)
        chk = ctk.CTkCheckBox(
            frame,
            text=label_text,
            fg_color=Colors.UI_BASIC,
            hover_color=Colors.adjust_brightness(Colors.UI_BASIC),
            checkmark_color=Colors.adjust_brightness(Colors.UI_BASIC, 1.8),
            variable=var,
        )
        chk.pack(side="left", padx=5)
        return var

    def _get_config_val(self, section, option):
        if self.config.has_option(section, option):
            return self.config.get(section, option)
        return ""

    def _save_config(self):
        """UIの値をconfigオブジェクトに書き戻してファイル保存"""
        try:
            # Paths
            self.config.set("Paths", "database_path", self.entry_db_path.get())
            self.config.set("Paths", "pdf_root_folder", self.entry_pdf_root.get())
            self.config.set("Paths", "pdf_archive_folder", self.entry_pdf_archive.get())
            self.config.set("Paths", "font_path", self.entry_font.get())

            # Automation
            self.config.set(
                "Automation", "enable_tesseract_ocr", str(self.var_ocr.get()).lower()
            )
            self.config.set(
                "Automation",
                "flatten_ink_annotations",
                str(self.var_flatten.get()).lower(),
            )
            self.config.set(
                "Automation",
                "auto_append_to_default_db",
                str(self.var_auto_db.get()).lower(),
            )
            self.config.set(
                "Automation",
                "create_individual_csv",
                str(self.var_export_csv.get()).lower(),
            )
            self.config.set(
                "Automation",
                "create_individual_csv",
                str(self.var_onthisday.get()).lower(),
            )

            # Watchdog (New)
            if not self.config.has_section("Watchdog"):
                self.config.add_section("Watchdog")

            self.config.set("Watchdog", "watch_dir", self.entry_watch_dir.get())
            self.config.set("Watchdog", "output_dir", self.entry_watch_output.get())
            self.config.set("Watchdog", "target_size", self.entry_watch_size.get())

            # ReportLab
            if not self.config.has_section("ReportLab"):
                self.config.add_section("ReportLab")

            self.config.set("ReportLab", "paper_size", self.entry_paper_size.get())
            self.config.set("ReportLab", "font", self.entry_rl_font.get())
            self.config.set("ReportLab", "author", self.entry_author.get())
            self.config.set("ReportLab", "title_prefix", self.entry_title_prefix.get())
            self.config.set("ReportLab", "max_pages", self.entry_max_pages.get())

            # 書き込み
            with open(self.config_path, "w", encoding="utf-8") as f:
                self.config.write(f)

            messagebox.showinfo(
                "保存完了",
                "設定を保存しました。\n変更を適用するにはアプリを再起動してください。",
            )
            self.destroy()

        except Exception as e:
            messagebox.showerror("保存エラー", f"設定の保存に失敗しました: {e}")


# 動作確認用
if __name__ == "__main__":
    root = ctk.CTk()
    root.configure(fg_color=(Colors.BACKGROUND_HOLLOW, Colors.BACKGROUND_DARK_HOLLOW))
    # テスト用ボタン
    btn = ctk.CTkButton(
        root,
        text="設定を開く",
        fg_color=Colors.UI_BASIC,
        hover_color=Colors.adjust_brightness(Colors.UI_BASIC),
        text_color="black",
        command=lambda: ConfigEditorWindow(root, "config.ini"),
    )
    btn.pack(pady=20)
    root.mainloop()

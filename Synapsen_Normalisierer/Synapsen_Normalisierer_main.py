import os
import sys
import shutil
import configparser
from tkinter import filedialog, messagebox
from pathlib import Path
import customtkinter as ctk
from dnd_window import DragAndDropWindow
from webclip_window import WebClipWindow

# PDF処理関数を別ファイルからインポート
from pdf_utils import (
    high_fidelity_flatten, normalize_pdf_to_papersize,
    embed_ocr_text_in_pdf,
    convert_image_to_pdf,
    convert_pil_image_to_pdf
)

A4_WIDTH = 595.276
A4_HEIGHT = 841.89
A5_WIDTH = 419.528
A5_HEIGHT = 595.276

SUPPORTED_EXTENSIONS = [".pdf", ".png", ".jpg", ".jpeg", ".bmp", ".tiff"]


class Synapsen_Normalisierer(ctk.CTk):
    """
    PDFのフォームフラット化とA4サイズ正規化を行うための
    CustomTkinterベースのGUIアプリケーション。

    Attributes:
        font_path (str): config.iniから読み込んだ、
                         フラット化時に使用するフォントファイルのパス。
        label (ctk.CTkLabel): アプリケーションのステータスを表示するラベル。
        run_button (ctk.CTkButton): 処理開始をトリガーするボタン。
    """

    def __init__(self):
        """
        アプリケーションウィンドウとウィジェットを初期化し、
        設定ファイルからフォントパスを読み込みます。
        """
        super().__init__()
        self.icon_path = self.get_icon_path()
        self.title("Synapsen Normalisierer")
        self.geometry("500x300")

        self.font_path = None
        self.paper_width = A4_WIDTH
        self.paper_height = A4_HEIGHT
        self.enable_tesseract_ocr = False
        self.config_data = {}  # 設定全体を保持する辞書
        self._load_config()

        # --- ウィジェットの配置 ---
        self.label = ctk.CTkLabel(
            self,
            text="フォームのテキスト化 及び 指定サイズ正規化を、\n注釈を維持したまま行います。"
        )
        self.label.pack(pady=20, padx=20)

        # 1. フォルダ処理モード
        self.folder_run_button = ctk.CTkButton(
            self,
            text="入力/出力フォルダを選んで処理実行",
            command=self.run_folder_process
        )
        self.folder_run_button.pack(pady=10, padx=10, fill="x", ipady=10)

        # 2. D&D / ペースト モード
        self.dnd_window_button = ctk.CTkButton(
            self,
            text="D&D / ペースト (個別ファイル) で正規化",
            command=self.open_dnd_window,
            fg_color="#585a9c",  # 桔梗色
            hover_color="#494B83"
        )
        self.dnd_window_button.pack(pady=10, padx=10, fill="x", ipady=10)

        # DNDウィンドウのインスタンスを保持
        self.dnd_window = None

        # 3. Webクリップ モード
        self.webclip_window_button = ctk.CTkButton(
            self,
            text="Webクリップ (URLからPDF化) で正規化",
            command=self.open_webclip_window,
            fg_color="#00695C",    # 濃い緑
            hover_color="#004D40"  # さらに濃い緑
        )
        self.webclip_window_button.pack(pady=10, padx=10, fill="x", ipady=10)

        # DNDウィンドウのインスタンスを保持
        self.dnd_window = None
        # WebClipウィンドウのインスタンスを保持
        self.webclip_window = None

        # 4. 共通ステータス表示ラベル
        self.status_label = ctk.CTkLabel(self, text="")
        self.status_label.pack(pady=(5, 10), padx=10)
        # --- [UI変更ここまで] ---

        # フォントパスの検証
        if not self.font_path or not Path(self.font_path).is_file():
            self.status_label.configure(
                text={
                    "エラー: config.iniで有効なフォントパスが指定されていません。" +
                    f"'{self.font_path}'"},
                text_color="orange"
            )
            self.folder_run_button.configure(state="disabled")
            self.dnd_window_button.configure(state="disabled")  # DNDボタンも無効化

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

    def _load_config(self) -> None:
        """
        config.iniファイルからフォントパスと用紙サイズを読み込みます。
        """
        # 0. config.ini のパスを決定
        if getattr(sys, 'frozen', False):
            base_path = os.path.dirname(sys.executable)
        else:
            base_path = os.path.dirname(os.path.abspath(__file__))

        if getattr(sys, 'frozen', False):
            config_path = os.path.join(base_path, 'config.ini')
        else:
            config_path = os.path.join(
                os.path.abspath(os.path.join(base_path, '..')), 'config.ini'
            )
        print(f"[DEBUG] Loading config from: {config_path}")

        config_dir = os.path.dirname(config_path)
        config = configparser.ConfigParser(interpolation=None)
        config.read(config_path, encoding='utf-8')

        # 1. フォントパスの読み込み
        font_path_from_config = config.get('Paths', 'font_path', fallback='')
        expanded_path = os.path.expandvars(font_path_from_config)  # 環境変数を展開

        if os.path.isabs(expanded_path):
            self.font_path = expanded_path
        else:
            self.font_path = os.path.join(config_dir, expanded_path)

        # config_data にも格納 (もし他のモジュールが参照する場合)
        self.config_data['font_path'] = self.font_path

        # 用紙サイズの読み込み
        paper_size_str = config.get(
            'LaTeX', 'paper_size', fallback='A4').upper()
        if paper_size_str == 'A5':
            self.paper_width = A5_WIDTH
            self.paper_height = A5_HEIGHT
        else:
            self.paper_width = A4_WIDTH
            self.paper_height = A4_HEIGHT

        self.config_data['paper_size'] = paper_size_str

        # 2. Tesseract OCR の有効/無効 設定
        self.enable_tesseract_ocr = config.getboolean(
            'Automation', 'enable_tesseract_ocr', fallback=False
        )
        self.config_data['enable_tesseract_ocr'] = self.enable_tesseract_ocr

        # 3. IndexKey の選択肢 (webclip_window が参照)
        keys_str = config.get('CommonplaceKeys', 'options', fallback='')
        self.config_data['commonplace_keys_options'] = [
            key.strip() for key in keys_str.split(',') if key.strip()
        ]

        # 4. KeyRect の座標 (webclip_window が参照)
        rect_str = config.get(
            'Extraction', 'key_rect', fallback='0,0,0,0').split(',')
        self.config_data['key_rect'] = tuple(map(float, rect_str))

        # 5. KeyIcons (webclip_window が参照)
        if config.has_section('KeyIcons'):
            self.config_data['key_icons'] = {
                k.lower(): v for k, v in config.items('KeyIcons')
            }
        else:
            self.config_data['key_icons'] = {}

        # 6. KeyColors (念のため読み込み)
        if config.has_section('KeyColors'):
            self.config_data['key_colors'] = {
                k.lower(): v for k, v in config.items('KeyColors')
            }
        else:
            self.config_data['key_colors'] = {}

    # --- DNDウィンドウを開く関数 ---
    def open_dnd_window(self):
        """
        「D&D / ペースト」ボタン押下時に、専用ウィンドウを開く。
        """
        # 既に開いている場合は、最前面に持ってくる
        if self.dnd_window is not None and self.dnd_window.winfo_exists():
            self.dnd_window.focus()
            self.dnd_window.grab_set()  # 再度モーダル化
        else:
            # self (メインアプリ自身) を親として渡す
            self.dnd_window = DragAndDropWindow(self)

    def open_webclip_window(self):
        """
        「Webクリップ」ボタン押下時に、専用ウィンドウを開く。
        """
        if (self.webclip_window is not None
                and self.webclip_window.winfo_exists()):
            self.webclip_window.focus()
            self.webclip_window.grab_set()
        else:
            self.webclip_window = WebClipWindow(self)

    def run_folder_process(self):
        """
        メイン機能：「フォルダ指定」で処理を実行する。
        """
        if not self.font_path or not Path(self.font_path).is_file():
            self.status_label.configure(
                text="エラー: config.iniで有効なフォントパスが指定されていません。",
                text_color="orange"
            )
            return

        source_folder = filedialog.askdirectory(title="入力元フォルダを選択してください")
        if not source_folder:
            return

        dest_folder = filedialog.askdirectory(title="出力先フォルダを選択してください")
        if not dest_folder:
            return

        if source_folder == dest_folder:
            messagebox.showerror("エラー", "入力元と出力先は異なるフォルダを選択してください。")
            return

        source_path = Path(source_folder)
        dest_path = Path(dest_folder)

        all_file_paths = []
        for ext in SUPPORTED_EXTENSIONS:
            all_file_paths.extend(source_path.glob(f"*{ext}"))
            if ext != ".pdf":
                all_file_paths.extend(source_path.glob(f"*{ext.upper()}"))

        # (Path, base_name) のタプルリストを作成
        items_to_process = [
            (p, p.stem) for p in sorted(list(set(all_file_paths)))]

        if not items_to_process:
            messagebox.showinfo(
                "情報",
                "処理対象のファイルが見つかりませんでした。\n" +
                f"(対象: {', '.join(SUPPORTED_EXTENSIONS)})"
                )
            self.status_label.configure(text="処理が完了しました（対象ファイルなし）。")
            return

        # 汎用処理関数を呼び出す
        self.execute_normalization_process(items_to_process, dest_path)

    def execute_normalization_process(
            self, items_to_process: list, dest_path: Path):
        """
        [共通処理関数]
        (Path, base_name) または (PIL.Image, base_name) のタプルリストを受け取り、
        {base_name}.pdf として正規化・出力する。
        """

        temp_dir = None

        try:
            all_items = sorted(items_to_process, key=lambda item: item[1])
            total_files = len(all_items)

            temp_dir = dest_path / "temp_flatten"
            temp_dir.mkdir(exist_ok=True)

            for i, (item_data, base_name) in enumerate(all_items):

                self.status_label.configure(
                    text=f"処理中 ({i+1}/{total_files}): {base_name}"
                )
                self.update_idletasks()

                # --- [ファイル名生成ロジック (base_name.pdf)] ---
                output_filename = f"{base_name}.pdf"
                final_output_pdf = dest_path / output_filename

                temp_flattened_pdf = temp_dir / f"flat_{base_name}.pdf"

                path_to_flatten: Path
                is_from_clipboard = False

                if isinstance(item_data, Path):
                    input_file_path = item_data
                else:
                    is_from_clipboard = True
                    input_file_path = temp_dir / f"{base_name}.pdf"
                    convert_pil_image_to_pdf(item_data, input_file_path)

                # --- [ステップ1: 画像 -> PDF] ---
                if (input_file_path.suffix.lower() != ".pdf"
                        and not is_from_clipboard):
                    self.status_label.configure(
                        text=f"({i+1}/{total_files}) 画像->PDF変換: {base_name}"
                    )
                    self.update_idletasks()

                    temp_converted_pdf = temp_dir / f"{base_name}.pdf"
                    try:
                        convert_image_to_pdf(
                            input_file_path, temp_converted_pdf)
                        path_to_flatten = temp_converted_pdf
                    except Exception as e:
                        print(f"警告: {base_name} のPDF変換に失敗: {e}")
                        continue
                else:
                    path_to_flatten = input_file_path

                # --- [ステップ2: フラット化] ---
                self.status_label.configure(
                    text=f"({i+1}/{total_files}) フラット化中: {base_name}"
                )
                self.update_idletasks()
                high_fidelity_flatten(
                    str(path_to_flatten),
                    str(temp_flattened_pdf),
                    self.font_path
                )

                # --- [ステップ3: 正規化] ---
                self.status_label.configure(
                    text=f"({i+1}/{total_files}) 正規化中: {base_name}"
                )
                self.update_idletasks()
                normalize_pdf_to_papersize(
                    str(temp_flattened_pdf),
                    str(final_output_pdf),
                    self.paper_width,
                    self.paper_height
                )

                # --- [ステップ4: OCR] ---
                self.status_label.configure(
                    text=f"({i+1}/{total_files}) OCR埋込処理中...: {base_name}"
                )
                self.update_idletasks()
                try:
                    embed_ocr_text_in_pdf(
                        str(final_output_pdf),
                        self.enable_tesseract_ocr,
                        self.font_path,
                        'jpn+jpn_vert'  # lang引数を明示
                    )
                except Exception as ocr_e:
                    messagebox.showerror("OCR エラー", str(ocr_e))
                    self.status_label.configure(text="OCRエラー。処理を中断しました。")
                    return

            messagebox.showinfo("完了", f"{total_files}個のPDF/画像ファイルの処理が完了しました。")
            self.status_label.configure(text="処理が完了しました。")

        except Exception as e:
            messagebox.showerror("エラー", f"処理中にエラーが発生しました:\n{e}")
            self.status_label.configure(text="エラーが発生しました。")

        finally:
            if temp_dir and temp_dir.exists():
                try:
                    shutil.rmtree(temp_dir)
                except Exception as e:
                    print(f"警告: 一時フォルダの削除に失敗しました: {e}")


if __name__ == "__main__":
    ctk.set_appearance_mode("System")
    app = Synapsen_Normalisierer()

    if app.icon_path:
        try:
            app.iconbitmap(default=str(app.icon_path))
        except Exception as e:
            print(f"Icon default setting error: {e}")
    else:
        print("警告: アイコンファイル (assets/synapsen.ico) が見つかりません。")

    app.mainloop()

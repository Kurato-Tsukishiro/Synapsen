import os
import sys
import shutil
import configparser
from tkinter import filedialog, messagebox
from pathlib import Path
import customtkinter as ctk

# PDF処理関数を別ファイルからインポート
from pdf_utils import (
    high_fidelity_flatten, normalize_pdf_to_papersize,
    embed_ocr_text_in_pdf,
    convert_image_to_pdf  # <--- [変更] 新しい関数をインポート
)


A4_WIDTH = 595.276
A4_HEIGHT = 841.89
A5_WIDTH = 419.528
A5_HEIGHT = 595.276

# <--- [追加] 処理対象の拡張子を定義
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
        self.geometry("500x250")

        self.font_path = None
        self.paper_width = A4_WIDTH  # デフォルト
        self.paper_height = A4_HEIGHT  # デフォルト
        self.enable_tesseract_ocr = False
        self._load_config()

        # --- ウィジェットの配置 ---
        self.label = ctk.CTkLabel(
            self,
            text="フォームのテキスト化 及び 指定サイズ正規化を、\n注釈を維持したまま行います。"
        )
        self.label.pack(pady=20, padx=20)

        self.run_button = ctk.CTkButton(
            self,
            text="処理を開始する",
            command=self.run_process
        )
        self.run_button.pack(pady=20, padx=20, ipady=10)

        # フォントパスの検証
        if not self.font_path or not Path(self.font_path).is_file():
            self.label.configure(
                text={
                    "エラー: config.iniで有効なフォントパスが指定されていません。" +
                    f"'{self.font_path}'"},
                text_color="orange"
            )
            self.run_button.configure(state="disabled")

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
        ( ... docstring ... )
        """
        # 0. config.ini のパスを決定
        if getattr(sys, 'frozen', False):
            # ( ... config_path 決定ロジック ... )
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

        # 用紙サイズの読み込み
        paper_size_str = config.get(
            'LaTeX', 'paper_size', fallback='A4').upper()
        if paper_size_str == 'A5':
            self.paper_width = A5_WIDTH
            self.paper_height = A5_HEIGHT
            print(f"[DEBUG] Paper size set to A5 ({
                self.paper_width}x{self.paper_height})")
        else:
            # デフォルトはA4
            self.paper_width = A4_WIDTH
            self.paper_height = A4_HEIGHT
            print(f"[DEBUG] Paper size set to A4 ({
                self.paper_width}x{self.paper_height})")

        # 2. Tesseract OCR の有効/無効 設定
        self.enable_tesseract_ocr = config.getboolean(
            'Automation', 'enable_tesseract_ocr', fallback=False
        )
        print(
            f"[DEBUG] Tesseract OCR (slow) enabled: {
                self.enable_tesseract_ocr}"
                )

    def run_process(self):
        """
        「処理を開始する」ボタン押下時のメイン処理。

        入力・出力フォルダをユーザーに選択させ、
        一時フォルダを作成し、対象のPDF「および画像」ファイル群に対して
        「画像->PDF変換」「フラット化」「正規化」「OCR」を順次実行します。
        """
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
        temp_dir = None  # finallyブロックで参照できるよう、外で定義

        try:
            # <--- [変更] サポートする全拡張子のファイルを検索 ---
            all_files = []
            for ext in SUPPORTED_EXTENSIONS:
                all_files.extend(source_path.glob(f"*{ext}"))
                # .pdf 以外は、大文字の拡張子 (e.g., .PNG) も検索
                if ext != ".pdf":
                    all_files.extend(source_path.glob(f"*{ext.upper()}"))

            # set() で重複を除去し、sorted() で処理順を一定にする
            all_files = sorted(list(set(all_files)))
            total_files = len(all_files)
            # <--- 変更ここまで ---

            if total_files == 0:
                messagebox.showinfo(
                    "情報",
                    "処理対象のファイルが見つかりませんでした。\n" +
                    f"(対象: {', '.join(SUPPORTED_EXTENSIONS)})"
                )
                self.label.configure(text="処理が完了しました（対象ファイルなし）。")
                return

            # 出力先フォルダ内に一時フォルダを作成
            temp_dir = dest_path / "temp_flatten"
            temp_dir.mkdir(exist_ok=True)

            for i, input_file in enumerate(all_files):
                self.label.configure(
                    text=f"処理中 ({i+1}/{total_files}): {input_file.name}"
                )
                self.update_idletasks()  # GUIの表示を強制更新

                # 最終的な出力ファイル名 (常時 .pdf)
                output_filename = input_file.with_suffix(".pdf").name
                final_output_pdf = dest_path / output_filename

                # フラット化処理の「出力先」
                temp_flattened_pdf = temp_dir / f"flat_{output_filename}"

                # フラット化処理の「入力元」 (PDFまたは変換後PDF)
                path_to_flatten: Path

                if input_file.suffix.lower() != ".pdf":
                    self.label.configure(
                        text=f"({i+1}/{total_files} )" +
                        f"画像->PDF変換: {input_file.name}"
                    )
                    self.update_idletasks()

                    # 変換したPDFを一時フォルダに保存
                    temp_converted_pdf = temp_dir / output_filename
                    try:
                        convert_image_to_pdf(input_file, temp_converted_pdf)
                        # 次のステップ（フラット化）の入力は、この変換したPDF
                        path_to_flatten = temp_converted_pdf
                    except Exception as e:
                        print(f"警告: {input_file.name} のPDF変換に失敗: {e}")
                        continue  # このファイルはスキップ
                else:
                    # 入力は元々PDF
                    path_to_flatten = input_file

                # 1. フォームをフラット化（一時フォルダに出力）
                self.label.configure(
                    text=f"({i+1}/{total_files}) フラット化中: {input_file.name}"
                )
                self.update_idletasks()
                high_fidelity_flatten(
                    str(path_to_flatten),
                    str(temp_flattened_pdf),
                    self.font_path
                )

                # 2. 指定サイズに正規化（最終出力先に出力）
                self.label.configure(
                    text=f"({i+1}/{total_files}) 正規化中: {input_file.name}"
                )
                self.update_idletasks()
                normalize_pdf_to_papersize(
                    str(temp_flattened_pdf),
                    str(final_output_pdf),
                    self.paper_width,
                    self.paper_height
                )

                # 3. OCR処理 (最終PDFに直接テキストを埋め込む)
                self.label.configure(
                    text={
                        f"({i+1}/{total_files}) OCR埋込処理中...: " +
                        f"{input_file.name}"
                        }
                )
                self.update_idletasks()
                try:
                    embed_ocr_text_in_pdf(
                        str(final_output_pdf),
                        self.enable_tesseract_ocr,
                        self.font_path
                    )
                except Exception as ocr_e:
                    # Tesseractが見つからない場合など
                    messagebox.showerror("OCR エラー", str(ocr_e))
                    self.label.configure(text="OCRエラー。処理を中断しました。")
                    return

            messagebox.showinfo("完了", f"{total_files}個のPDF/画像ファイルの処理が完了しました。")
            self.label.configure(text="処理が完了しました。")

        except Exception as e:
            messagebox.showerror("エラー", f"処理中にエラーが発生しました:\n{e}")
            self.label.configure(text="エラーが発生しました。")

        finally:
            # 最後に必ず一時フォルダを削除する
            if temp_dir and temp_dir.exists():
                try:
                    shutil.rmtree(temp_dir)
                except Exception as e:
                    print(f"警告: 一時フォルダの削除に失敗しました: {e}")


if __name__ == "__main__":
    ctk.set_appearance_mode("System")
    app = Synapsen_Normalisierer()
# 1. 実行ファイル(.exe)かスクリプト(.py)かによって基準パスを取得
    if getattr(sys, 'frozen', False):
        # .exe実行の場合（実行ファイルの場所）
        base_path = os.path.dirname(sys.executable)

        # .exe の場合: 'assets\synapsen.ico' (base_path と同じ階層)
        icon_path = os.path.join(base_path, "assets", "synapsen.ico")
    else:
        # スクリプト実行の場合（.pyファイルの場所）
        base_path = os.path.dirname(os.path.abspath(__file__))

        # スクリプトの場合: '..\assets\synapsen.ico' (base_path の1つ上の階層)
        icon_path = os.path.join(base_path, "..", "assets", "synapsen.ico")

    # 3. アイコンを設定 (存在する場合のみ)
    # os.path.normpath() は '..' を解決してきれいなパスにします
    iconfile = os.path.normpath(icon_path)

    if app.icon_path:  # <-- クラス内で取得したパスを利用
        try:
            # 'default=' を指定し、OSダイアログ(エクスプローラ等)にも適用
            app.iconbitmap(default=str(app.icon_path))
        except Exception as e:
            print(f"Icon default setting error: {e}")
    else:
        print("警告: アイコンファイル (assets/synapsen.ico) が見つかりません。")
    app.mainloop()

import os
import sys
import shutil
import configparser
from tkinter import filedialog, messagebox
from pathlib import Path
import customtkinter as ctk

# --- ローカルモジュールのインポート ---
# ロギング設定
import logging

# サブウィンドウ
from dnd_window import DragAndDropWindow
from webclip_window import WebClipWindow

# PDF処理バックエンド
from pdf_utils import (
    high_fidelity_flatten, normalize_pdf_to_papersize,
    embed_ocr_text_in_pdf,
    convert_image_to_pdf,
    convert_pil_image_to_pdf,
    convert_markdown_to_pdf
)

# --- 定数 ---
A4_WIDTH = 595.276
A4_HEIGHT = 841.89
A5_WIDTH = 419.528
A5_HEIGHT = 595.276

SUPPORTED_EXTENSIONS = [".pdf", ".png", ".jpg", ".jpeg", ".bmp", ".tiff"]

# ==============================================================================
# ロギング設定の初期化
# ==============================================================================
# 親ディレクトリ(ルート)をパスに追加して logging_setup.py をインポート可能にする
current_dir = Path(__file__).parent
root_dir = current_dir.parent
if str(root_dir) not in sys.path:
    sys.path.append(str(root_dir))

try:
    from logging_setup import setup_logging
    # アプリ名を指定して初期化
    setup_logging("Synapsen_Normalisierer")
    logger = logging.getLogger("Normalisierer")  # このファイル用のロガー取得
except ImportError:
    # logging_setup.py がない場合のフォールバック（print出力）
    print("Warning: logging_setup.py not found. Logging disabled.")

    class MockLogger:
        def info(self, msg): print(f"[INFO] {msg}")

        def error(self, msg, exc_info=None):
            print(f"[ERROR] {msg} {exc_info if exc_info else ''}")

        def warning(self, msg): print(f"[WARN] {msg}")

    logger = MockLogger()
# ==============================================================================


class Synapsen_Normalisierer(ctk.CTk):
    """
    PDFのフォームフラット化と指定サイズ正規化を行う
    CustomTkinterベースのGUIアプリケーション。

    以下の機能を提供します:
    - フォルダ単位での一括正規化。
    - ドラッグ＆ドロップおよびペーストによる個別ファイルの正規化。
    - URLを指定したWebクリップと正規化。

    Attributes:
        icon_path (Path | None): アプリケーションアイコンのパス。
        font_path (str | None): config.iniから読み込んだフォントパス。
        paper_width (float): 正規化後の用紙幅 (ポイント)。
        paper_height (float): 正規化後の用紙高 (ポイント)。
        enable_tesseract_ocr (bool): Tesseract OCRを実行するか否か。
        config_data (dict): config.iniから読み込んだ設定の辞書。
        dnd_window (DragAndDropWindow | None): D&Dウィンドウのインスタンス。
        webclip_window (WebClipWindow | None): WebClipウィンドウのインスタンス。
    """

    def __init__(self):
        """
        アプリケーションウィンドウとウィジェットを初期化し、
        設定ファイル (config.ini) を読み込みます。
        """
        super().__init__()
        self.icon_path = self.get_icon_path()
        self.title("Synapsen Normalisierer")
        self.geometry("500x300")

        # --- 設定値の初期化 ---
        self.font_path = None
        self.paper_width = A4_WIDTH
        self.paper_height = A4_HEIGHT
        self.enable_tesseract_ocr = False
        self.config_data = {}  # WebClipウィンドウなどが参照する設定全体
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

        self.dnd_window = None  # DNDウィンドウのインスタンスを保持

        # 3. Webクリップ モード
        self.webclip_window_button = ctk.CTkButton(
            self,
            text="Webクリップ (URLからPDF化) で正規化",
            command=self.open_webclip_window,
            fg_color="#00695C",    # 濃い緑
            hover_color="#004D40"  # さらに濃い緑
        )
        self.webclip_window_button.pack(pady=10, padx=10, fill="x", ipady=10)

        self.webclip_window = None  # WebClipウィンドウのインスタンスを保持

        # 4. 共通ステータス表示ラベル
        self.status_label = ctk.CTkLabel(self, text="")
        self.status_label.pack(pady=(5, 10), padx=10)

        # --- 設定の検証とUIの初期化 ---
        self._validate_config_and_update_ui()

    def get_icon_path(self) -> Path | None:
        """
        実行環境(.exe or .py)に応じて、
        プロジェクトルートの 'assets' フォルダにある
        'synapsen.ico' のパスを返します。

        Returns:
            Path | None: アイコンファイルへのPathオブジェクト。見つからない場合はNone。
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
            logger.error(f"Error finding icon path: {e}")
        return None

    def _load_config(self) -> None:
        """
        config.iniファイルから各種設定を読み込み、クラス属性にセットします。

        読み込む設定:
        - Paths: font_path, tags_data_path, database_path
        - LaTeX: paper_size, font, author, title_prefix
        - Automation: enable_tesseract_ocr, auto_append_to_default_db, ...
        - Extraction: key_rect
        - CommonplaceKeys: options
        - KeyIcons, KeyColors
        """
        try:
            # 0. config.ini のパスを決定
            if getattr(sys, 'frozen', False):
                base_path = os.path.dirname(sys.executable)
            else:
                base_path = os.path.dirname(os.path.abspath(__file__))

            if getattr(sys, 'frozen', False):
                config_path = os.path.join(base_path, 'config.ini')
            else:
                config_path = os.path.join(
                    os.path.abspath(os.path.join(base_path, '..')),
                    'config.ini'
                )

            if not os.path.exists(config_path):
                messagebox.showerror(
                    "設定エラー", f"config.ini が見つかりません。\nパス: {config_path}")
                self.font_path = None
                return

            logger.debug(f"Loading config from: {config_path}")

            config_dir = os.path.dirname(config_path)
            config = configparser.ConfigParser(interpolation=None)
            config.read(config_path, encoding='utf-8')

            # 1. フォントパスの読み込み (Normalisierer の中核機能)
            font_path_from_config = config.get(
                'Paths', 'font_path', fallback='')
            expanded_path = os.path.expandvars(font_path_from_config)
            if os.path.isabs(expanded_path):
                self.font_path = expanded_path
            else:
                self.font_path = os.path.join(config_dir, expanded_path)
            self.config_data['font_path'] = self.font_path

            # 2. 用紙サイズの読み込み
            paper_size_str = config.get(
                'LaTeX', 'paper_size', fallback='A4').upper()
            if paper_size_str == 'A5':
                self.paper_width = A5_WIDTH
                self.paper_height = A5_HEIGHT
            else:
                self.paper_width = A4_WIDTH
                self.paper_height = A4_HEIGHT
            self.config_data['paper_size'] = paper_size_str

            # 3. Tesseract OCR の有効/無効 設定
            self.enable_tesseract_ocr = config.getboolean(
                'Automation', 'enable_tesseract_ocr', fallback=False
            )
            self.config_data[
                'enable_tesseract_ocr'] = self.enable_tesseract_ocr

            # 3.5. LaTeXフォント名の読み込み (Pandoc用)
            self.config_data['latex_font'] = config.get(
                'LaTeX', 'font', fallback='MS UI Gothic'
            )

            # 4. WebClipウィンドウが参照するその他の設定
            keys_str = config.get('CommonplaceKeys', 'options', fallback='')
            self.config_data['commonplace_keys_options'] = [
                key.strip() for key in keys_str.split(',') if key.strip()
            ]

            rect_str = config.get(
                'Extraction', 'key_rect', fallback='0,0,0,0').split(',')
            self.config_data['key_rect'] = tuple(map(float, rect_str))

            if config.has_section('KeyIcons'):
                self.config_data['key_icons'] = {
                    k.lower(): v for k, v in config.items('KeyIcons')
                }
            else:
                self.config_data['key_icons'] = {}

            if config.has_section('KeyColors'):
                self.config_data['key_colors'] = {
                    k.lower(): v for k, v in config.items('KeyColors')
                }
            else:
                self.config_data['key_colors'] = {}

        except Exception as e:
            messagebox.showerror("設定エラー", f"config.ini の読み込みに失敗しました:\n{e}")
            self.font_path = None  # エラー時はフォントパスをNoneに

    def _validate_config_and_update_ui(self) -> None:
        """
        読み込まれた設定（特にフォントパス）を検証し、
        UIの状態（ボタンの有効/無効、ステータスラベル）を更新します。
        """
        if not self.font_path or not Path(self.font_path).is_file():
            self.status_label.configure(
                text=(
                    "エラー: config.iniで有効なフォントパスが指定されていません。\n" +
                    f"'{self.font_path}'"
                    ),
                text_color="orange"
            )
            self.folder_run_button.configure(state="disabled")
            self.dnd_window_button.configure(state="disabled")
            self.webclip_window_button.configure(state="disabled")
        else:
            # フォントパスが有効なら、全ボタンを有効化
            self.status_label.configure(text="設定を読み込みました。")
            self.folder_run_button.configure(state="normal")
            self.dnd_window_button.configure(state="normal")
            self.webclip_window_button.configure(state="normal")

    def open_dnd_window(self) -> None:
        """
        「D&D / ペースト」ボタン押下時に、DragAndDropWindow を開きます。
        既にウィンドウが存在する場合は、それを最前面に表示します。
        """
        if self.dnd_window is not None and self.dnd_window.winfo_exists():
            self.dnd_window.focus()
            self.dnd_window.grab_set()
        else:
            # self (メインアプリ自身) を親として渡す
            self.dnd_window = DragAndDropWindow(self)

    def open_webclip_window(self) -> None:
        """
        「Webクリップ」ボタン押下時に、WebClipWindow を開きます。
        既にウィンドウが存在する場合は、それを最前面に表示します。
        """
        if (self.webclip_window is not None
                and self.webclip_window.winfo_exists()):
            self.webclip_window.focus()
            self.webclip_window.grab_set()
        else:
            self.webclip_window = WebClipWindow(self)

    def run_folder_process(self) -> None:
        """
        「フォルダ指定」で処理を実行します。
        入力元・出力先フォルダをユーザーに選択させ、
        対象ファイルを収集して `execute_normalization_process` を呼び出します。
        """
        # フォントパスの再検証 (config.iniが後から変更される可能性)
        if not self.font_path or not Path(self.font_path).is_file():
            self._validate_config_and_update_ui()
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

        # サポート対象の拡張子を持つファイルを取得
        all_file_paths = []
        for ext in SUPPORTED_EXTENSIONS:
            all_file_paths.extend(source_path.glob(f"*{ext}"))
            # 大文字の拡張子も検索 (例: .PNG)
            if ext != ".pdf":
                all_file_paths.extend(source_path.glob(f"*{ext.upper()}"))

        # (Pathオブジェクト, 拡張子なしのファイル名) のタプルリストを作成
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
            self, items_to_process: list, dest_path: Path) -> None:
        """
        [共通処理関数]
        正規化処理の本体。
        入力アイテムのリストを受け取り、指定された出力先パスに
        正規化されたPDF (`{base_name}.pdf`) を出力します。

        処理ステップ:
        1-A.(Markdownの場合) PandocでPDFに変換
        1-B. (画像/PILの場合) PDFに変換
        2. high_fidelity_flatten (フォームのテキスト化)
        3. normalize_pdf_to_papersize (指定サイズに中央配置)
        4. embed_ocr_text_in_pdf (OCRテキストレイヤーの埋め込み)

        Args:
            items_to_process (list): 処理対象のアイテムのリスト。
                各アイテムは以下のタプル形式:
                - (Path, str): (入力ファイルのPathオブジェクト, 出力ベース名)
                - (PIL.Image.Image, str): (Pillowイメージ, 出力ベース名)
            dest_path (Path): 出力先フォルダのPathオブジェクト。
        """
        temp_dir = None
        try:
            # アイテムを出力ベース名でソート
            all_items = sorted(items_to_process, key=lambda item: item[1])
            total_files = len(all_items)

            # 一時フォルダを出力先フォルダ内に作成
            temp_dir = dest_path / "temp_flatten"
            temp_dir.mkdir(exist_ok=True)

            paper_size_str = self.config_data.get('paper_size', 'A4')

            for i, (item_data, base_name) in enumerate(all_items):

                status_prefix = f"処理中 ({i+1}/{total_files}):"
                self.status_label.configure(
                    text=f"{status_prefix} {base_name}"
                )
                self.update_idletasks()

                # --- [ファイル名生成ロジック] ---
                output_filename = f"{base_name}.pdf"
                final_output_pdf = dest_path / output_filename
                # フラット化後の一時ファイルパス
                temp_flattened_pdf = temp_dir / f"flat_{base_name}.pdf"

                path_to_flatten: Path  # フラット化対象のPDFパス

                # --- [ステップ 1-A: MD -> PDF 変換] ---
                if (isinstance(item_data, Path) and
                        item_data.suffix.lower()) == ".md":
                    self.status_label.configure(
                        text=f"{status_prefix} MD->PDF変換: {base_name}"
                    )
                    self.update_idletasks()

                    # 一時フォルダに {base_name}.pdf として変換
                    temp_converted_md_pdf = temp_dir / f"md_{base_name}.pdf"
                    try:
                        convert_markdown_to_pdf(
                            item_data,
                            temp_converted_md_pdf,
                            paper_size_str
                        )
                        # 変換後のPDFを、次のパイプラインの入力 (item_data) として上書き
                        item_data = temp_converted_md_pdf
                    except Exception as e:
                        logger.warning(f"警告: {base_name} のMarkdown変換に失敗: {e}")
                        # pandoc がない場合など
                        messagebox.showerror(
                            "Markdown変換エラー",
                            f"{base_name} の変換に失敗しました:\n{e}",
                            parent=self
                        )
                        continue  # このファイルはスキップ

                # --- [ステップ1-B: 画像 -> PDF変換] ---
                if isinstance(item_data, Path):
                    input_file_path = item_data

                    # 入力アイテムがPathだが、PDFでない (画像) 場合
                    if input_file_path.suffix.lower() != ".pdf":
                        self.status_label.configure(
                            text=f"{status_prefix} 画像->PDF変換: {base_name}"
                        )
                        self.update_idletasks()

                        # 一時フォルダに {base_name}.pdf として変換
                        temp_converted_pdf = temp_dir / f"{base_name}.pdf"
                        try:
                            convert_image_to_pdf(
                                input_file_path, temp_converted_pdf)
                            path_to_flatten = temp_converted_pdf
                        except Exception as e:
                            logger.warning(f"警告: {base_name} のPDF変換に失敗: {e}")
                            continue  # このファイルはスキップ
                    else:
                        # 入力アイテムがPDF (またはMDから変換されたPDF) の場合
                        path_to_flatten = input_file_path

                else:
                    # 入力アイテムがPathでない (PIL.Image) 場合
                    temp_converted_pdf = temp_dir / f"{base_name}.pdf"
                    try:
                        convert_pil_image_to_pdf(item_data, temp_converted_pdf)
                        path_to_flatten = temp_converted_pdf
                    except Exception as e:
                        logger.warning(
                            f"警告: {base_name} (クリップボード) のPDF変換に失敗: {e}")
                        continue  # このアイテムはスキップ

                # --- [ステップ2: フラット化 (フォームのテキスト化)] ---
                self.status_label.configure(
                    text=f"{status_prefix} フラット化中: {base_name}"
                )
                self.update_idletasks()
                high_fidelity_flatten(
                    str(path_to_flatten),
                    str(temp_flattened_pdf),
                    self.font_path
                )

                # --- [ステップ3: 正規化 (サイズ統一)] ---
                self.status_label.configure(
                    text=f"{status_prefix} 正規化中: {base_name}"
                )
                self.update_idletasks()
                normalize_pdf_to_papersize(
                    str(temp_flattened_pdf),
                    str(final_output_pdf),
                    self.paper_width,
                    self.paper_height
                )

                # --- [ステップ4: OCR埋め込み] ---
                self.status_label.configure(
                    text=f"{status_prefix} OCR埋込処理中...: {base_name}"
                )
                self.update_idletasks()

                # embed_ocr_text_in_pdf は final_output_pdf を直接上書き変更する
                embed_ocr_text_in_pdf(
                    str(final_output_pdf),
                    self.enable_tesseract_ocr,
                    self.font_path,
                    'jpn+jpn_vert'
                )

            messagebox.showinfo(
                "完了", f"{total_files}個のPDF/画像/MDファイルの処理が完了しました。")
            self.status_label.configure(text="処理が完了しました。")

        except Exception as e:
            # TesseractNotFoundError や Pandoc がない場合のエラーもここで捕捉
            messagebox.showerror("エラー", f"処理中にエラーが発生しました:\n{e}")
            self.status_label.configure(text=f"エラーが発生しました: {e}")

        finally:
            # 正常終了・異常終了に関わらず、一時フォルダを削除
            if temp_dir and temp_dir.exists():
                try:
                    shutil.rmtree(temp_dir)
                except Exception as e:
                    logger.warning(f"一時フォルダの削除に失敗しました: {e}")


if __name__ == "__main__":
    ctk.set_appearance_mode("System")
    app = Synapsen_Normalisierer()

    if app.icon_path:
        try:
            # 'default=' を指定し、OSダイアログ(エクスプローラ等)にも適用
            app.iconbitmap(default=str(app.icon_path))
        except Exception as e:
            logger.error(f"Icon default setting error: {e}")
    else:
        logger.warning("警告: アイコンファイル (assets/synapsen.ico) が見つかりません。")

    app.mainloop()

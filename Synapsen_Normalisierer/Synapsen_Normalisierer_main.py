from enum import Enum
import os
import sys
import shutil
import configparser
from tkinter import filedialog, messagebox
from pathlib import Path
import unicodedata
import customtkinter as ctk
import re

# --- ローカルモジュールのインポート ---
# ロギング設定
import logging

# サブウィンドウ
from dnd_window import DragAndDropWindow
from webclip_window import WebClipWindow

# PDF処理バックエンド
from pdf_utils import (
    OCREngineType,
    high_fidelity_flatten,
    normalize_pdf_to_papersize,
    embed_ocr_text_in_pdf,
    convert_image_to_pdf,
    convert_pil_image_to_pdf,
    convert_document_to_pdf,
    embed_normalization_metadata,
    transplant_ocr_text_from_pdf,
    run_yomitoku_ocr,
)

current_dir = Path(__file__).parent
root_dir = current_dir.parent
if str(root_dir) not in sys.path:
    sys.path.append(str(root_dir))

from theme import SemanticColors as Colors  # noqa: E402

# --- 定数 ---
A4_WIDTH = 595.276
A4_HEIGHT = 841.89
A5_WIDTH = 419.528
A5_HEIGHT = 595.276

SUPPORTED_EXTENSIONS = [
    ".pdf",
    ".png",
    ".jpg",
    ".jpeg",
    ".bmp",
    ".tiff",
    ".md",
    ".txt",
    ".docx",
    ".rtf",
    ".odt",
]

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
        def info(self, msg):
            print(f"[INFO] {msg}")

        def error(self, msg, exc_info=None):
            print(f"[ERROR] {msg} {exc_info if exc_info else ''}")

        def warning(self, msg):
            print(f"[WARN] {msg}")

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
        tags_data_path (str | None): config.iniから読み込んだ事前定義タグを保存しているテキストファイルのパス。
        paper_width (float): 正規化後の用紙幅 (ポイント)。
        paper_height (float): 正規化後の用紙高 (ポイント)。
        ocr_engine (OCREngineType): OCRのエンジン設定。
        use_existing_text (bool):
            YomiTokuによるOCR化時、既存テキストを利用できるか(全体設定)。
        cp_keys_list (list[str]): 登録されているCommonplaceKeyのリスト。
        cp_keys_re_pattern (re.Pattern):
            登録されているCommonplaceKeyのコンパイル済みの検索パターンオブジェクト。
        flatten_ink (bool): インク注釈をフラット化するか否か。
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
        self.configure(
            fg_color=(Colors.BACKGROUND_HOLLOW, Colors.BACKGROUND_DARK_HOLLOW)
        )
        self.geometry("500x300")

        # --- 設定値の初期化 ---
        self.font_path = None
        self.tags_data_path = None
        self.paper_width = A4_WIDTH
        self.paper_height = A4_HEIGHT
        self.ocr_engine = OCREngineType.NONE
        self.use_existing_text = False
        self.cp_keys_list = None
        self.cp_keys_re_pattern = None
        self.flatten_ink = True
        self.config_data = {}  # WebClipウィンドウなどが参照する設定全体
        self._load_config()

        # --- ウィジェットの配置 ---
        self.label = ctk.CTkLabel(
            self,
            text="フォームのテキスト化 及び 指定サイズ正規化を、\n注釈を維持したまま行います。",
        )
        self.label.pack(pady=20, padx=20)

        # 1. フォルダ処理モード
        self.folder_run_button = ctk.CTkButton(
            self,
            text="入力/出力フォルダを選んで処理実行",
            command=self.run_folder_process,
            fg_color=Colors.UI_BASIC,
            hover_color=Colors.adjust_brightness(Colors.UI_BASIC),
            text_color="black",
        )
        self.folder_run_button.pack(pady=10, padx=10, fill="x", ipady=10)

        # 2. D&D / ペースト モード
        self.dnd_window_button = ctk.CTkButton(
            self,
            text="D&D / ペースト (個別ファイル) で正規化",
            command=self.open_dnd_window,
            fg_color=Colors.UI_SECONDARY,
            hover_color=Colors.adjust_brightness(Colors.UI_SECONDARY),
        )
        self.dnd_window_button.pack(pady=10, padx=10, fill="x", ipady=10)

        self.dnd_window = None  # DNDウィンドウのインスタンスを保持

        # 3. Webクリップ モード
        self.webclip_window_button = ctk.CTkButton(
            self,
            text="Webクリップ (URLからPDF化) で正規化",
            command=self.open_webclip_window,
            fg_color=Colors.UI_EXPORT,
            hover_color=Colors.adjust_brightness(Colors.UI_EXPORT, 0.6),
        )
        self.webclip_window_button.pack(pady=10, padx=10, fill="x", ipady=10)

        self.webclip_window = None  # WebClipウィンドウのインスタンスを保持

        # 4. 共通ステータス表示ラベル
        self.status_label = ctk.CTkLabel(self, text="")
        self.status_label.pack(pady=(5, 10), padx=10)

        # --- PDFTagsの読み込み ---
        self.tags_map = self._load_pdf_tags()

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
            if getattr(sys, "frozen", False):
                # .exe実行の場合 (exeと同じフォルダがプロジェクトルート)
                project_root = Path(sys.executable).parent
            else:
                # .pyスクリプト実行の場合 (このファイルの親フォルダがプロジェクトルート)
                project_root = Path(__file__).parent.parent

            icon_path = project_root / "assets" / "synapsen.ico"

            if icon_path.is_file():
                return icon_path
        except Exception as e:
            logger.error(f"Error finding icon path: {e}")
        return None

    def _load_config(self) -> None:
        """
        config.iniファイルから各種設定を読み込み、クラス属性にセットします。

        読み込む設定:
        - Paths: font_path, tags_data_path
        - ReportLab: paper_size, font, author, title_prefix
        - Automation:
            ocr_engine,
            use_existing_text,
            auto_append_to_default_db,
            flatten_ink_annotations
        - Extraction: key_rect
        - CommonplaceKeys: options
        - KeyIcons, KeyColors
        """
        try:
            # 0. config.ini のパスを決定
            if getattr(sys, "frozen", False):
                base_path = os.path.dirname(sys.executable)
            else:
                base_path = os.path.dirname(os.path.abspath(__file__))

            if getattr(sys, "frozen", False):
                config_path = os.path.join(base_path, "config.ini")
            else:
                config_path = os.path.join(
                    os.path.abspath(os.path.join(base_path, "..")), "config.ini"
                )

            if not os.path.exists(config_path):
                messagebox.showerror(
                    "設定エラー", f"config.ini が見つかりません。\nパス: {config_path}"
                )
                self.font_path = None
                return

            logger.debug(f"Loading config from: {config_path}")

            config_dir = os.path.dirname(config_path)
            config = configparser.ConfigParser(interpolation=None)
            config.read(config_path, encoding="utf-8")

            # 1. パスの読み込み
            # 1-I. フォントパスの読み込み (Normalisierer の中核機能)
            font_path_from_config = config.get("Paths", "font_path", fallback="")
            expanded_font_path = os.path.expandvars(font_path_from_config)
            if os.path.isabs(expanded_font_path):
                self.font_path = expanded_font_path
            else:
                self.font_path = os.path.join(config_dir, expanded_font_path)
            self.config_data["font_path"] = self.font_path

            # 1-II. 事前定義タグ設定ファイルのパスの読み込み
            tags_data_path_from_config = config.get(
                "Paths", "tags_data_path", fallback=""
            )
            expanded_tags_data_path = os.path.expandvars(tags_data_path_from_config)
            if os.path.isabs(expanded_tags_data_path):
                self.tags_data_path = expanded_tags_data_path
            else:
                self.tags_data_path = os.path.join(config_dir, expanded_tags_data_path)
            self.config_data["tags_data_path"] = self.tags_data_path

            # --- セクション名の切り替え処理 ---

            # 2. 用紙サイズの読み込み
            paper_size_str = config.get(
                "ReportLab", "paper_size", fallback="A4"
            ).upper()
            if paper_size_str == "A5":
                self.paper_width = A5_WIDTH
                self.paper_height = A5_HEIGHT
            else:
                self.paper_width = A4_WIDTH
                self.paper_height = A4_HEIGHT
            self.config_data["paper_size"] = paper_size_str

            # 3. OCR関連設定の取得
            # 3-I. エンジン設定
            ocr_engine_str = config.get("Automation", "use_ocr_engine", fallback="")
            self.ocr_engine = OCREngineType(
                ocr_engine_str
            )  # 文字列で指定されたエンジン名を列挙型に変換して取得
            self.config_data["use_ocr_engine"] = self.ocr_engine

            # 3-II ollamaの設定
            self.ollama_model = config.get(
                "Automation", "ollama_ocr_model", fallback=""
            )
            self.ollama_api_url = config.get(
                "Automation",
                "ollama_api_url",
                fallback="http://localhost:11434/api/generate",
            )

            self.config_data["ollama_model"] = self.ollama_model
            self.config_data["ollama_api_url"] = self.ollama_api_url

            # 3-III. 既存テキストの文字カウント判定からコモンプレイスキーを除外する為の設定取得
            cp_keys_str = config.get("CommonplaceKeys", "options", fallback="")
            self.cp_keys_list = [
                opt.strip() for opt in cp_keys_str.split(",") if opt.strip()
            ]
            self.cp_keys_re_pattern = self._get_CommonplaceKeys_pattern(
                self.cp_keys_list
            )
            # config_dataへの記録は7番で実行

            # 3-IV. YomiTokuによるOCR化時、既存テキスト利用可能にするかの設定取得
            self.use_existing_text = config.getboolean(
                "Automation", "use_existing_text", fallback=False
            )
            self.config_data["use_existing_text"] = self.use_existing_text

            # 4. インク注釈のフラット化設定
            self.flatten_ink = config.getboolean(
                "Automation", "flatten_ink_annotations", fallback=True
            )
            self.config_data["flatten_ink_annotations"] = self.flatten_ink

            # 5. フォント設定 (現在未使用だが、設定としては読み込んでおく)
            self.config_data["reportlab_font"] = config.get(
                "ReportLab", "font", fallback=""
            )

            # 6. 引用Key用QRコードのサイズの読み込み
            self.config_data["refs_qr_size"] = config.getint(
                "Extraction", "refs_qr_size", fallback=75
            )

            # 7. WebClipウィンドウが参照するその他の設定
            keys_str = config.get("CommonplaceKeys", "options", fallback="")
            self.config_data["commonplace_keys_options"] = [
                key.strip() for key in keys_str.split(",") if key.strip()
            ]

            rect_str = config.get("Extraction", "key_rect", fallback="0,0,0,0").split(
                ","
            )
            self.config_data["key_rect"] = tuple(map(float, rect_str))

            if config.has_section("KeyIcons"):
                self.config_data["key_icons"] = {
                    k.lower(): v for k, v in config.items("KeyIcons")
                }
            else:
                self.config_data["key_icons"] = {}

            if config.has_section("KeyColors"):
                self.config_data["key_colors"] = {
                    k.lower(): v for k, v in config.items("KeyColors")
                }
            else:
                self.config_data["key_colors"] = {}

        except Exception as e:
            messagebox.showerror(
                "設定エラー", f"config.ini の読み込みに失敗しました:\n{e}"
            )
            self.font_path = None

    def _validate_config_and_update_ui(self) -> None:
        """
        読み込まれた設定（特にフォントパス）を検証し、
        UIの状態（ボタンの有効/無効、ステータスラベル）を更新します。
        """
        if not self.font_path or not Path(self.font_path).is_file():
            self.status_label.configure(
                text=(
                    "エラー: config.iniで有効なフォントパスが指定されていません。\n"
                    + f"'{self.font_path}'"
                ),
                text_color=Colors.LABEL_WARNING,
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
        if self.webclip_window is not None and self.webclip_window.winfo_exists():
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

        source_folder = filedialog.askdirectory(
            title="入力元フォルダを選択してください"
        )
        if not source_folder:
            return

        dest_folder = filedialog.askdirectory(title="出力先フォルダを選択してください")
        if not dest_folder:
            return

        if source_folder == dest_folder:
            messagebox.showerror(
                "エラー", "入力元と出力先は異なるフォルダを選択してください。"
            )
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
        items_to_process = [(p, p.stem) for p in sorted(list(set(all_file_paths)))]

        if not items_to_process:
            messagebox.showinfo(
                "情報",
                "処理対象のファイルが見つかりませんでした。\n"
                + f"(対象: {', '.join(SUPPORTED_EXTENSIONS)})",
            )
            self.status_label.configure(text="処理が完了しました（対象ファイルなし）。")
            return

        # 汎用処理関数を呼び出す
        self.execute_normalization_process(items_to_process, dest_path)

    class IndividualControlUsingFileName:
        """ファイル名を使用した、処理制御の状態を管理するクラス

        Attributes:
            remove_pattern (Flag_Type): ファイル名からフラグとなる文字列を省くための正規表現パターン。
            use_ollama (bool): OCRエンジンとしてollamが指定されているか。
            use_yomitoku_force (bool):
                YomiTokuによるOCR処理時、既存テキストの利用をファイル名により無効化されているか。
                (全体設定に関わらず、YomiTokuによるOCR処理を既存テキストの有無を無視して行うか。)
        """

        class Flag_Type(Enum):
            """ファイル名によって、指定する処理の種類"""

            LLM_OCR = "LLM_OCR"
            YOMITOKU_OCR = "YOMITOKU_OCR"

        NAME_FLAG_DICT = {
            "_hand": Flag_Type.LLM_OCR,
            "_llm": Flag_Type.LLM_OCR,
            "_ytocr": Flag_Type.YOMITOKU_OCR,
        }
        """ファイル名の接尾辞と接尾辞に対応する処理の種類の組み合わせ"""

        _REMOVE_PATTERN = re.compile(
            "|".join(map(re.escape, NAME_FLAG_DICT.keys())), re.IGNORECASE
        )
        """ファイル名からフラグを削除する為の正規表現パターン"""

        def __init__(self, file_name: str):
            self.remove_pattern = self._REMOVE_PATTERN
            self.use_ollama = self._get_use_ollama(file_name)
            self.use_yomitoku_force = self._get_use_yomitoku_force(file_name)

        @classmethod
        def check_file_name_flag(_cls, flag: Flag_Type, file_name: str) -> bool:
            """指定した処理の実行が、ファイル名によって有効になっているか

            Args:
                flag (Flag_Type): 指定する処理
                file_name (str): 指定の処理が有効であるか、確認する対象のファイル名

            Returns:
                bool: True: 有効 / False: 無効
            """
            if not isinstance(file_name, str):
                return False

            flag_list = [k for k, v in _cls.NAME_FLAG_DICT.items() if v == flag]
            name_flag = any(item in file_name.lower() for item in flag_list)

            return name_flag

        @classmethod
        def _get_use_ollama(_cls, file_name: str):
            use_ollama = _cls.check_file_name_flag(_cls.Flag_Type.LLM_OCR, file_name)
            if use_ollama:
                logger.info("ollamaフラグを検知しました。")
            return use_ollama

        @classmethod
        def _get_use_yomitoku_force(_cls, file_name: str):
            use_yomitoku_force = _cls.check_file_name_flag(
                _cls.Flag_Type.YOMITOKU_OCR, file_name
            )
            if use_yomitoku_force:
                logger.info("YomiTokuによるOCR処理を優先するフラグを検知しました。")
            return use_yomitoku_force

    def execute_normalization_process(
        self, items_to_process: list, dest_path: Path
    ) -> None:
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

            # configから用紙サイズ設定を取得 ("A4" or "A5")
            paper_size_str = self.config_data.get("paper_size", "A4")

            for i, (item_data, base_name) in enumerate(all_items):
                # OCR用PDFファイル(_ocr.pdf) 自身の処理をスキップ
                # ループの最初で判定し、以降の不要な処理(正規表現など)を回避する
                original_item_path = (
                    item_data if isinstance(item_data, Path) else Path(str(item_data))
                )
                if original_item_path.name.lower().endswith("_ocr.pdf"):
                    logger.info(
                        f"外部OCRファイル ({original_item_path.name}) のため、単独の正規化処理をスキップします。",
                        extra={"sensitive": True},
                    )
                    if hasattr(self, "status_label"):
                        self.status_label.configure(
                            text=f"スキップ (OCRファイル): {original_item_path.name}"
                        )
                        self.update_idletasks()
                    continue

                status_prefix = f"処理中 ({i+1}/{total_files}):"
                self.status_label.configure(text=f"{status_prefix} {base_name}")
                self.update_idletasks()

                # --- [ファイル名生成ロジック変更] ---
                # 1. OCRモードの判定 (先に判定しておく: 元のbase_nameを使用)
                file_name_control_state = self.IndividualControlUsingFileName(base_name)

                # 2. 出力ファイル名からフラグを削除 (base_name -> output_base_name)
                output_base_name = file_name_control_state.remove_pattern.sub(
                    "", base_name
                )

                output_filename = f"{output_base_name}.pdf"
                final_output_pdf = dest_path / output_filename

                # フラット化後の一時ファイルパス
                temp_flattened_pdf = temp_dir / f"flat_{base_name}.pdf"

                path_to_flatten: Path  # フラット化対象のPDFパス

                # --- [ステップ 1-A: ドキュメント形式 (MD, TXT, DOCX等) -> PDF 変換] ---
                if isinstance(item_data, Path) and item_data.suffix.lower() not in [
                    ".pdf",
                    ".png",
                    ".jpg",
                    ".jpeg",
                    ".bmp",
                    ".tiff",
                ]:

                    self.status_label.configure(
                        text=f"{status_prefix} ドキュメント->PDF変換: {base_name}"
                    )
                    self.update_idletasks()

                    # 一時フォルダに変換
                    temp_converted_doc_pdf = temp_dir / f"doc_{base_name}.pdf"
                    try:
                        # 汎用関数 convert_document_to_pdf を使用
                        convert_document_to_pdf(
                            item_data, temp_converted_doc_pdf, paper_size_str
                        )
                        # 変換後のPDFを、次の処理（フラット化・正規化）の入力として上書き
                        item_data = temp_converted_doc_pdf

                    except Exception as e:
                        logger.warning(
                            f"警告: {base_name} のドキュメント変換に失敗: {e}"
                        )
                        messagebox.showerror(
                            "ドキュメント変換エラー",
                            f"{base_name} の変換に失敗しました:\n{e}",
                            parent=self,
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
                            convert_image_to_pdf(input_file_path, temp_converted_pdf)
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
                            f"警告: {base_name} (クリップボード) のPDF変換に失敗: {e}"
                        )
                        continue  # このアイテムはスキップ

                # --- [ステップ2: フラット化 (フォームのテキスト化 + インク焼き込み)] ---
                self.status_label.configure(
                    text=f"{status_prefix} フラット化中: {base_name}"
                )
                self.update_idletasks()
                high_fidelity_flatten(
                    str(path_to_flatten),
                    str(temp_flattened_pdf),
                    self.font_path,
                    flatten_ink=self.flatten_ink,
                )

                # --- 3: OCR (正規化前 実行) ---
                # --- [ステップ3-I: OCR用PDFファイルからのOCRテキスト移植] ---
                # 入力元がファイルである場合、同階層の _ocr.pdf を探す
                original_item_path = item_data if isinstance(item_data, Path) else None
                ocr_transplanted = False  # 正規化前にOCR処理を実行しているか
                is_ocr_file_exists = False  # OCR用PDFが存在するか

                if original_item_path:
                    ocr_source_path = (
                        original_item_path.parent / f"{original_item_path.stem}_ocr.pdf"
                    )
                    # OCR用PDFの有無を記録
                    is_ocr_file_exists = (
                        ocr_source_path.exists() and ocr_source_path.is_file()
                    )
                    if is_ocr_file_exists:
                        self.status_label.configure(
                            text=(
                                f"{status_prefix} 透明化テキスト統合中: "
                                f"{ocr_source_path.name}"
                            )
                        )
                        self.update_idletasks()
                        try:
                            # 座標ズレを防ぐため、正規化(リサイズ)前の temp_flattened_pdf にテキストを移植する
                            ocr_transplanted = transplant_ocr_text_from_pdf(
                                str(temp_flattened_pdf),
                                str(ocr_source_path),
                                self.font_path,
                            )
                        except Exception as e:
                            logger.error(f"透明化テキスト統合エラー: {e}")

                # --- [ステップ3-II: YomiTokuによるOCR化を実行] ---
                # 正規化後のOCRではレイアウトの崩れが発生したため、それを防ぐ為に正規化前に実行している
                if self.ocr_engine == OCREngineType.YomiToku and not is_ocr_file_exists:
                    try:
                        # 既存テキストの使用が可能か（Trueの場合、OCRをスキップしてテキストを再利用）
                        should_use_existing_text = (
                            self.use_existing_text
                            and not file_name_control_state.use_yomitoku_force
                        )

                        ocr_transplanted = run_yomitoku_ocr(
                            temp_flattened_pdf,
                            self.font_path,
                            should_use_existing_text,
                            self.cp_keys_re_pattern,
                        )

                        if not ocr_transplanted:
                            self.ocr_engine = OCREngineType.tesseract
                            logger.warning(
                                "YomiTokuによるOCR処理に失敗したため、エンジンを tesseract に変更しました。"
                            )
                    except Exception as e:
                        self.ocr_engine = OCREngineType.tesseract
                        logger.error(
                            f"YomiTokuによるOCR化に失敗したため、エンジンを tesseract に変更しました。{e}"
                        )

                # --- [ステップ4: 正規化 (サイズ統一)] ---
                self.status_label.configure(
                    text=f"{status_prefix} 正規化中: {base_name}"
                )
                self.update_idletasks()
                normalize_pdf_to_papersize(
                    str(temp_flattened_pdf),
                    str(final_output_pdf),
                    self.paper_width,
                    self.paper_height,
                    target_format=paper_size_str,
                )

                # --- 5: OCR (正規化後 実行) ---
                # OCR用PDFから移植した場合やYomiTokuでOCR化した場合は、処理をスキップする。
                if not ocr_transplanted:
                    # 移植していない場合は処理を実行する。
                    self._process_ocr_sequence(
                        base_name,
                        final_output_pdf,
                        file_name_control_state.use_ollama,
                        status_prefix,
                    )

                # --- [ステップ6: 正規化メタデータ (フラグ・タグ) 埋め込み] ---
                # ファイル名からタグを抽出
                extracted_tags = []
                for match in re.finditer(r"\[(.*?)\]", base_name):
                    raw_tags = [t.strip() for t in match.group(1).split(",")]
                    for t in raw_tags:
                        t_lower = t.lower()
                        if t_lower in self.tags_map:
                            # アンダースコアでタグを分割する(階層構造の展開)
                            parts = self.tags_map[t_lower].split("_")
                            for i in range(len(parts)):
                                hierarchical_tag = "_".join(parts[: i + 1])
                                if hierarchical_tag not in extracted_tags:
                                    extracted_tags.append(hierarchical_tag)

                self.status_label.configure(
                    text=f"{status_prefix} メタデータ埋込: {output_base_name}"
                )
                self.update_idletasks()

                # 包括的なメタデータ埋め込み関数を呼び出し
                embed_normalization_metadata(
                    str(final_output_pdf), extra_tags=extracted_tags
                )

            messagebox.showinfo(
                "完了", f"{total_files}個のPDF/画像/MDファイルの処理が完了しました。"
            )
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

    def _process_ocr_sequence(
        self,
        base_name: str,
        ocr_target_pdf_path_str: str,
        use_ollama: bool,
        status_prefix: str,
    ) -> None:
        """
        UIの更新からOCRの実行までの一連のシーケンスを制御する中継関数
        (dnd_windowにも同様の関数がある為、これを変更した時はそちら側の変更も行う)

        Args:
            base_name (str): OCR処理対象の元々のファイル名
            ocr_target_pdf_path_str (str): OCR処理対象のPDFのパス
            use_ollama (bool): ollamaをOCR処理に使用するか(ファイル名によるフラグが指定されていたか)
            status_prefix (str): 「何ファイル中、何ファイル目のOCR処理を実行しているか」のラベル表示用テキスト
        """

        def _determine_ocr_engine(use_ollama: bool) -> tuple[bool, OCREngineType]:
            """
            OCR処理の実行が可能であるかの判定と、使用するエンジン名の取得を行う。

            Args:
                use_ollama (bool): ollamaをOCR処理に使用するか(ファイル名によるフラグが指定されていたか)

            Returns:
                tuple[bool, str]: 実行可能かどうかの真偽値、及び使用するエンジン名。
            """
            if use_ollama:
                if self.ollama_model:  # config上で、モデル設定が行われているか
                    return True, "ollama"
                else:
                    # ログ出力のみ
                    logger.warning(
                        "Ollamaフラグを検出しましたが、configにモデル設定がないためOCRをスキップします。"
                    )
                    return False, "NONE"

            # Python 3.10以降の match-case 構文で分岐
            match self.ocr_engine:
                case OCREngineType.tesseract:
                    return True, self.ocr_engine
                case OCREngineType.YomiToku | OCREngineType.NONE:
                    return False, self.ocr_engine
                case _:
                    # 上記のどれにも一致しない場合（不正な設定）
                    logger.warning(
                        f"OCRのエンジンの設定が不正である為、OCRをスキップします。({self.ocr_engine.value})"
                    )
                    return False, OCREngineType.NONE

        def _embed_ocr_text(
            ocr_target_pdf_path_str: str,
            engine: OCREngineType,
        ) -> None:
            """OCR処理の呼び出し

            Args:
                ocr_target_pdf_path_str (str): OCR処理対象のPDFのパス
                engine (OCREngineType): OCR処理を実行するエンジン
            """

            # OCRを行うかはチェック済みなので、ここでは純粋にOCR処理の呼び出しを行うだけ
            embed_ocr_text_in_pdf(
                pdf_path_str=str(ocr_target_pdf_path_str),
                enable_ocr=True,  # ここにこれたなら必ずTrue
                font_path=self.font_path,
                cp_keys_re_pattern=self.cp_keys_re_pattern,
                ocr_engine=engine,  # 使用可否確認済みのエンジン
                ollama_config={
                    "model": self.ollama_model,
                    "url": self.ollama_api_url,
                },
            )

        # 1. 状態の判定ロジックを呼び出す（UI操作なし）
        should_run, engine = _determine_ocr_engine(use_ollama)

        # 2. ガード句：実行不要及び実行不能なら、画面を汚さずにこの関数を即座に抜ける
        if not should_run:
            return

        # 3. 実行確定時のみUIを更新する
        self.status_label.configure(
            text=f"{status_prefix} OCR処理中({engine.value})...: {base_name}",
            extra={"sensitive": True},
        )
        self.update_idletasks()

        # 4. 実際のOCR（embed_ocr_text_in_pdf の呼び出し）を実行
        _embed_ocr_text(ocr_target_pdf_path_str=ocr_target_pdf_path_str, engine=engine)

    @staticmethod
    def _get_CommonplaceKeys_pattern(cp_keys_list: list[str]) -> re.Pattern:
        """CommonplaceKeyの正規表現パターンを取得する。

        Args:
            cp_keys_list (list[str]): 登録されているCommonplaceKeyのリスト

        Returns:
            str: CommonplaceKeyの正規表現パターン
        """
        # 文字コード正規化
        cleaned_keys = [
            unicodedata.normalize("NFKC", opt.strip()) for opt in cp_keys_list
        ]
        ck_pattern_str = "|".join(map(re.escape, cleaned_keys))
        ck_pattern = re.compile(ck_pattern_str)

        return ck_pattern

    def _load_pdf_tags(self) -> dict[str, str]:
        """
        PDFTags.txt を読み込み、エイリアスを含むタグマッピング辞書を作成します。
        キーは小文字化されたタグ名・エイリアス、値は正式なタグ名です。
        """
        tags_map = {}
        try:
            tags_file_path = Path(self.tags_data_path)

            if not tags_file_path.exists():
                logger.warning(f"PDFTags.txt が見つかりません: {tags_file_path}")
                return tags_map

            with open(tags_file_path, "r", encoding="utf-8") as f:
                for line in f:
                    line = line.strip()
                    if not line or line.startswith("#"):
                        continue

                    # 正式タグ名 [エイリアス1, エイリアス2...] の形式をパース
                    match = re.match(r"^([^\[\s]+)(?:\s*\[(.*?)\])?\s*$", line)
                    if match:
                        main_tag = match.group(1).strip()
                        # 正式タグ名自身もマップに登録
                        tags_map[main_tag.lower()] = main_tag

                        aliases_str = match.group(2)
                        if aliases_str:
                            for alias in aliases_str.split(","):
                                alias = alias.strip()
                                if alias:
                                    tags_map[alias.lower()] = main_tag
                    else:
                        main_tag = line.split()[0].strip()
                        tags_map[main_tag.lower()] = main_tag

        except Exception as e:
            logger.error(f"PDFTags.txt の読み込み中にエラーが発生しました: {e}")

        return tags_map


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
        logger.warning(
            "警告: アイコンファイル (assets/synapsen.ico) が見つかりません。"
        )

    app.mainloop()

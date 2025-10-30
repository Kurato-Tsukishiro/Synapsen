import os
import sys
import shutil
import configparser
from tkinter import filedialog, messagebox
from pathlib import Path
import customtkinter as ctk

# PDF処理関数を別ファイルからインポート
from pdf_utils import high_fidelity_flatten, normalize_pdf_to_papersize

A4_WIDTH = 595.276
A4_HEIGHT = 841.89
A5_WIDTH = 419.528
A5_HEIGHT = 595.276


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

    def run_process(self):
        """
        「処理を開始する」ボタン押下時のメイン処理。

        入力・出力フォルダをユーザーに選択させ、
        一時フォルダを作成し、対象のPDFファイル群に対して
        「フラット化」と「正規化」を順次実行します。
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
            pdf_files = list(source_path.glob("*.pdf"))
            total_files = len(pdf_files)

            if total_files == 0:
                messagebox.showinfo("情報", "処理対象のPDFファイルが見つかりませんでした。")
                self.label.configure(text="処理が完了しました（対象ファイルなし）。")
                return

            # 出力先フォルダ内に一時フォルダを作成
            temp_dir = dest_path / "temp_flatten"
            temp_dir.mkdir(exist_ok=True)

            for i, pdf_file in enumerate(pdf_files):
                self.label.configure(
                    text=f"処理中 ({i+1}/{total_files}): {pdf_file.name}"
                    )
                self.update_idletasks()  # GUIの表示を強制更新

                temp_flattened_pdf = temp_dir / pdf_file.name
                final_output_pdf = dest_path / pdf_file.name

                # 1. フォームをフラット化（一時フォルダに出力）
                high_fidelity_flatten(
                    str(pdf_file),
                    str(temp_flattened_pdf),
                    self.font_path
                )

                # 2. 指定サイズに正規化（最終出力先に出力）
                normalize_pdf_to_papersize(
                    str(temp_flattened_pdf),
                    str(final_output_pdf),
                    self.paper_width,
                    self.paper_height
                )

            messagebox.showinfo("完了", f"{total_files}個のPDFファイルの処理が完了しました。")
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

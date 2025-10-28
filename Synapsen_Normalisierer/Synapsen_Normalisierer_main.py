import os
import sys
import shutil
import configparser
from tkinter import filedialog, messagebox
from pathlib import Path
import customtkinter as ctk

# PDF処理関数を別ファイルからインポート
from pdf_utils import high_fidelity_flatten, normalize_pdf_to_a4


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
        self.title("Synapsen Normalisierer")
        self.geometry("500x250")

        self.font_path = self._load_font_path()

        # --- ウィジェットの配置 ---
        self.label = ctk.CTkLabel(
            self,
            text="フォームのテキスト化 及び A4縦サイズ正規化を、\n注釈を維持したまま行います。 "
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

    def _load_font_path(self) -> str:
        """
        config.iniファイルからフォントパスを読み込みます。

        実行可能ファイル(PyInstaller等)かスクリプト実行かを判別し、
        `config.ini`の相対パスを解決します。
        環境変数（%LOCALAPPDATA%など）も展開します。

        Returns:
            str: 'Paths'セクションの'font_path'の値。見つからない場合は空文字。
        """
        if getattr(sys, 'frozen', False):
            # 実行可能ファイルの場合 (e.g., PyInstaller)
            base_path = os.path.dirname(sys.executable)
        else:
            # 通常のスクリプト実行の場合
            base_path = os.path.dirname(os.path.abspath(__file__))

        # .exe実行かスクリプト実行かで config.ini の場所を切り替える
        if getattr(sys, 'frozen', False):
            # .exe実行の場合（config.ini は .exe と同じフォルダ）
            config_path = os.path.join(base_path, 'config.ini')
        else:
            # スクリプト実行の場合（config.ini は .py の1つ上のフォルダ）
            config_path = os.path.join(
                os.path.abspath(os.path.join(base_path, '..')), 'config.ini'
            )
        print(f"[DEBUG] Loading config from: {config_path}")

        # config.ini があるフォルダのパスを基準として定義
        config_dir = os.path.dirname(config_path)

        config = configparser.ConfigParser(interpolation=None)
        config.read(config_path, encoding='utf-8')

        font_path_from_config = config.get('Paths', 'font_path', fallback='')
        expanded_path = os.path.expandvars(font_path_from_config)  # 環境変数を展開

        if os.path.isabs(expanded_path):
            # configの値が絶対パス（または環境変数展開後、絶対パスになった）の場合
            return expanded_path
        else:
            # configの値が相対パスの場合、config_dir と結合する
            return os.path.join(config_dir, expanded_path)

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

                # 2. A4サイズに正規化（最終出力先に出力）
                normalize_pdf_to_a4(
                    str(temp_flattened_pdf),
                    str(final_output_pdf)
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
    app.mainloop()

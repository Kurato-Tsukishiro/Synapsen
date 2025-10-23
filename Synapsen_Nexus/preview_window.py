import customtkinter as ctk
# utilsからメモ欄構築関数をインポート
from utils import build_memo_display


class NotePreviewWindow(ctk.CTkToplevel):
    """
    ノートのメタデータをプレビュー表示するための専用Toplevelウィンドウ。

    メインウィンドウから独立して表示され、ノートの詳細とPDFへの
    ショートカットを提供する。
    """
    def __init__(self, parent_app, note_data):
        """
        NotePreviewWindowを初期化する。

        Args:
            parent_app (DigitalCommonplaceBook):
                このウィンドウを呼び出したメインアプリケーションのインスタンス。
                (self.parent_app.df や self.parent_app.open_preview_window の
                 呼び出しに使用)
            note_data (pd.Series):
                表示するノートのデータ（DataFrameの1行）。
        """
        super().__init__(parent_app)
        self.parent_app = parent_app  # メインアプリ本体
        self.note_data = note_data

        title = self.note_data.get('title', 'N/A')
        self.title(f"プレビュー: {title}")
        self.geometry("450x450")
        self.transient(parent_app)  # 常にメインウィンドウより手前に表示
        self.grab_set()  # このウィンドウを閉じるまでメインを操作不可にする

        self.grid_columnconfigure(1, weight=1)
        self.grid_rowconfigure(4, weight=1)  # メモ欄が伸縮するように

        # --- ウィジェットの作成 ---

        # 1. タイトル
        ctk.CTkLabel(
            self, text="タイトル:", anchor="w"
            ).grid(row=0, column=0, padx=10, pady=5, sticky="w")
        ctk.CTkLabel(
            self, text=title, wraplength=300, justify="left", anchor="w"
            ).grid(row=0, column=1, padx=10, pady=5, sticky="ew")

        # 2. キー
        ctk.CTkLabel(
            self, text="キー:", anchor="w"
            ).grid(row=1, column=0, padx=10, pady=5, sticky="w")
        ctk.CTkLabel(
            self, text=self.note_data.get('key', ''), anchor="w"
            ).grid(row=1, column=1, padx=10, pady=5, sticky="ew")

        # 3. Index Key
        ctk.CTkLabel(
            self, text="インデックス キー:", anchor="w"
            ).grid(row=2, column=0, padx=10, pady=5, sticky="w")
        ctk.CTkLabel(
            self, text=self.note_data.get('commonplace_key', ''), anchor="w"
            ).grid(row=2, column=1, padx=10, pady=5, sticky="ew")

        # 4. タグ
        ctk.CTkLabel(
            self, text="タグ:", anchor="w"
            ).grid(row=3, column=0, padx=10, pady=5, sticky="w")
        tags_str = str(self.note_data.get('tags', '')).replace(';', ', ')
        ctk.CTkLabel(
            self, text=tags_str, wraplength=300, justify="left", anchor="w"
            ).grid(row=3, column=1, padx=10, pady=5, sticky="ew")

        # 5. メモ (読み取り専用)
        ctk.CTkLabel(
            self, text="メモ:", anchor="w"
            ).grid(row=4, column=0, padx=10, pady=5, sticky="nw")
        self.memo_display_frame = ctk.CTkScrollableFrame(self)
        self.memo_display_frame.grid(
            row=4, column=1, padx=10, pady=5, sticky="nsew"
            )

        # メモ欄の動的ビルドメソッドを呼び出す
        self._build_memo_display()

        # 6. 「PDFを開く」ボタン
        pdf_button = ctk.CTkButton(
            self, text="PDFを開く", command=self.open_pdf_action
            )
        pdf_button.grid(row=5, column=0, columnspan=2, padx=10, pady=10)

    def open_pdf_action(self):
        """「PDFを開く」ボタンが押されたときの処理。"""
        # メインアプリのopen_pdfメソッドを呼び出す
        self.parent_app.open_pdf(self.note_data)
        self.destroy()  # PDFを開いたらプレビューは閉じる

    def _build_memo_display(self):
        """
        プレビューウィンドウのメモ欄に、クリック可能なリンク付きラベルを生成する。
        utils.build_memo_display を使用する。
        """
        memo_text = str(self.note_data.get('memo', ''))

        # プレビューウィンドウの幅に合わせてテキストが折り返すように設定
        frame_width = 300

        # メインアプリのDataFrameとプレビュー展開メソッドをコールバックとして渡す
        build_memo_display(
            self.memo_display_frame,
            memo_text,
            self.parent_app.df,  # リンク先タイトルの検索用
            self.parent_app.open_preview_window,  # リンククリック時の動作
            frame_width
        )

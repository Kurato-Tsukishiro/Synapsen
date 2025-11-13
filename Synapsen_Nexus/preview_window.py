import customtkinter as ctk

# utilsからメモ欄構築関数をインポート
from utils import (
    build_memo_display, find_backlinks_df, build_references_display,
    get_pdf_page_image  # <-- [追加] PDF画像化ヘルパーをインポート
)


# ==============================================================================
# 簡易プレビューウィンドウ (読み取り専用)
# ==============================================================================
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

        # CTkImageオブジェクトへの参照を保持 (ガベージコレクション対策)
        self.preview_image_object = None

        self._custom_icon_path = None  # 強制設定するアイコンパス
        if hasattr(parent_app, 'icon_path') and parent_app.icon_path:
            self._custom_icon_path = str(parent_app.icon_path)

            if self._custom_icon_path:
                try:
                    super().iconbitmap(self._custom_icon_path)
                except Exception as e:
                    print(f"Initial icon set error: {e}")

        title = self.note_data.get('title', 'N/A')
        self.title(f"プレビュー: {title}")

        # 縦サイズを 600 -> 750 に拡大
        self.geometry("450x750")
        self.transient(parent_app)  # 常にメインウィンドウより手前に表示

        self.grid_columnconfigure(1, weight=1)

        # [変更] グリッドの重み設定 (プレビュー、メモ、引用元)
        self.grid_rowconfigure(4, weight=1)  # <-- ★ 4. PDFプレビュー
        self.grid_rowconfigure(6, weight=1)  # <-- ★ 6. メモ欄
        self.grid_rowconfigure(8, weight=1)  # <-- ★ 8. 引用元欄

        # --- ウィジェットの作成 (すべて読み取り専用) ---

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

        tags_str = str(self.note_data.get('tags', ''))
        tags_list = [tag for tag in tags_str.split(';') if tag]
        tags_display = ", ".join(tags_list)

        ctk.CTkLabel(
            self, text=tags_display, wraplength=300, justify="left", anchor="w"
            ).grid(row=3, column=1, padx=10, pady=5, sticky="ew")

        # 4. PDFプレビュー
        self.pdf_preview_label = ctk.CTkLabel(
            self,
            text="プレビューを読込中...",
            fg_color="gray20",
            anchor="center",
            text_color="gray70"
        )
        self.pdf_preview_label.grid(
            row=4, column=0, columnspan=2, padx=10, pady=10, sticky="nsew"
        )

        # 5. メモ (ラベル)
        ctk.CTkLabel(
            self, text="メモ:", anchor="w"
            ).grid(row=5, column=0, padx=10, pady=5, sticky="nw")

        # 6. メモ (フレーム)
        self.memo_display_frame = ctk.CTkScrollableFrame(self)
        self.memo_display_frame.grid(
            row=6, column=1, padx=10, pady=5, sticky="nsew"
            )
        self._build_memo_display()

        # 7. 引用元 (ラベル)
        ctk.CTkLabel(
            self, text="引用元:", anchor="w"
            ).grid(row=7, column=0, padx=10, pady=5, sticky="nw")

        # 8. 引用元 (フレーム)
        self.references_display_frame = ctk.CTkScrollableFrame(
            self, label_text="このノートを引用"
            )
        self.references_display_frame.grid(
            row=8, column=1, padx=10, pady=5, sticky="nsew"
            )
        self._build_references_display()

        # 9. ボタンエリア (PDFを開く + 編集)
        button_frame = ctk.CTkFrame(self, fg_color="transparent")
        button_frame.grid(row=9, column=0, columnspan=2, padx=10, pady=10)

        # PDFを開くボタン
        pdf_button = ctk.CTkButton(
            button_frame, text="PDFを開く", command=self.open_pdf_action
        )
        pdf_button.pack(side="left", padx=10)

        # 編集ボタン
        edit_button = ctk.CTkButton(
            button_frame,
            text="編集する",
            command=self.edit_note_action,
            fg_color="#585a9c",
            hover_color="#494B83"
        )
        edit_button.pack(side="left", padx=10)

        # --- PDFプレビューの読み込み処理 ▼ ---
        max_preview_width = 250  # プレビュー表示の最大幅
        pil_image = get_pdf_page_image(
            self.note_data,
            self.parent_app.loaded_db_path,
            self.parent_app.pdf_root_folder,
            max_width=max_preview_width
        )

        if pil_image:
            self.preview_image_object = ctk.CTkImage(
                light_image=pil_image,
                dark_image=pil_image,
                size=(pil_image.width, pil_image.height)
            )
            self.pdf_preview_label.configure(
                image=self.preview_image_object, text="",
                fg_color="transparent"
            )
        else:
            self.pdf_preview_label.configure(
                text="プレビューの読み込みに失敗しました", text_color="#D9534F"
            )

        self.grab_set()  # このウィンドウを閉じるまでメインを操作不可にする

    def iconbitmap(self, *args, **kwargs):
        """
        iconbitmap の呼び出しをインターセプト（横取り）する。

        CustomTkinterが内部でこのメソッドを呼び出して
        アイコンをデフォルトに戻そうとしても、
        強制的にカスタムアイコンを設定し直す。
        """
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

    def open_pdf_action(self):
        """「PDFを開く」ボタンが押されたときの処理。"""
        # メインアプリのopen_pdfメソッドを呼び出す
        self.parent_app.open_pdf(self.note_data)
        self.destroy()  # PDFを開いたらプレビューは閉じる

    def edit_note_action(self):
        """
        編集ボタンが押されたときの処理。
        プレビューウィンドウを閉じ、メインアプリの編集ダイアログを開く。
        """
        # 編集後にプレビューの内容が古くなるのを防ぐため、先にウィンドウを閉じます
        self.destroy()
        # メインアプリ(Synapsen_Nexus)の open_edit_dialog を呼び出します
        self.parent_app.open_edit_dialog(self.note_data)

    def _build_memo_display(self):
        """
        プレビューウィンドウのメモ欄に、クリック可能なリンク付きラベルを生成する。
        """
        memo_text = str(self.note_data.get('memo', ''))
        frame_width = 300  # プレビューウィンドウの幅

        # メインアプリのDataFrameと「プレビュー展開」メソッドをコールバックとして渡す
        build_memo_display(
            self.memo_display_frame,
            memo_text,
            self.parent_app.df,  # リンク先タイトルの検索用
            self.parent_app.open_preview_window,  # リンククリック時の動作
            frame_width
        )

    def _build_references_display(self):
        """
        プレビューウィンドウの引用元欄を構築する。
        """
        current_key = self.note_data.get('key', '')

        # メインアプリのDataFrameと設定を使って検索
        backlinks_df = find_backlinks_df(
            self.parent_app.df, current_key
        )

        # utilsの関数を使って引用元UIを構築
        build_references_display(
            self.references_display_frame,
            backlinks_df,
            self.parent_app.open_preview_window,  # Callback to main app
            self.parent_app.key_icons,
            self.parent_app.key_colors
        )

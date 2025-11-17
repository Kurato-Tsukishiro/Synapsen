import customtkinter as ctk
import pandas as pd
import fitz

from utils import (
    build_memo_display, build_references_display,
    get_pdf_document_for_note,
    get_pdf_page_image_from_doc
)

import logging
logger = logging.getLogger(__name__)


# ==============================================================================
# 簡易プレビューウィンドウ (読み取り専用)
# ==============================================================================
class NotePreviewWindow(ctk.CTkToplevel):
    """
    ノートのメタデータをプレビュー表示するための専用Toplevelウィンドウ。

    メインウィンドウから独立して表示され、ノートの詳細とPDFへの
    ショートカットを提供する。
    """
    def __init__(self, parent_app, note_data, default_view_mode='compact'):
        """
        NotePreviewWindowを初期化する。

        Args:
            parent_app (DigitalCommonplaceBook):
                このウィンドウを呼び出したメインアプリケーションのインスタンス。
                (self.parent_app.df や self.parent_app.open_preview_window の
                 呼び出しに使用)
            note_data (pd.Series):
                表示するノートのデータ（DataFrameの1行）。
            default_view_mode (str, optional):
                縮小表示(リンクからの呼び出し)か拡大表示(メインウィンドウからの詳細表示)か
        """
        super().__init__(parent_app)
        self.parent_app = parent_app  # メインアプリ本体
        self.note_data = note_data

        # CTkImageオブジェクトへの参照を保持 (ガベージコレクション対策)
        self.preview_image_object = None

        # ページめくり用の状態変数
        self.current_pdf_doc: fitz.Document | None = None  # PDFドキュメント本体
        self.current_pdf_page_index = 0          # 現在表示中のページ (0-indexed)
        self.current_pdf_page_count = 0          # このノートの総ページ数
        self.current_pdf_doc_start_index = 0     # doc内の開始インデックス

        # プレビューウィンドウのレイアウト状態とサイズを管理
        self.is_compact_view = (default_view_mode == 'compact')
        self.compact_geometry = "450x750"  # 垂直レイアウト    (リンク呼び出し)
        self.full_geometry = "1350x750"    # 2カラムレイアウト (本体呼び出し)

        self._custom_icon_path = None
        if hasattr(parent_app, 'icon_path') and parent_app.icon_path:
            self._custom_icon_path = str(parent_app.icon_path)

            if self._custom_icon_path:
                try:
                    super().iconbitmap(self._custom_icon_path)
                except Exception as e:
                    logger.error(f"Initial icon set error: {e}")

        title = self.note_data.get('title', 'N/A')
        self.title(f"プレビュー: {title}")
        self.transient(parent_app)

        self.protocol("WM_DELETE_WINDOW", self.on_close)

        # --- [Col 0] PDFプレビューコンテナ ---
        self.pdf_preview_container = ctk.CTkFrame(self)
        self.pdf_preview_container.grid_rowconfigure(1, weight=1)
        self.pdf_preview_container.grid_columnconfigure(0, weight=1)

        # PDFプレビュー用の上部バー (ラベル + ページめくり)
        pdf_top_bar = ctk.CTkFrame(
            self.pdf_preview_container, fg_color="transparent"
        )
        pdf_top_bar.grid(
            row=0, column=0, padx=5, pady=5, sticky="ew"
        )
        ctk.CTkLabel(
            pdf_top_bar, text="PDFプレビュー:", anchor="w"
        ).pack(side="left", padx=(0, 10))

        # ページめくりボタン (初期状態は無効)
        self.pdf_prev_button = ctk.CTkButton(
            pdf_top_bar, text="<", width=30, state="disabled",
            command=self.show_prev_page
        )
        self.pdf_prev_button.pack(side="left", padx=5)

        self.pdf_page_label = ctk.CTkLabel(pdf_top_bar, text="(-/-)", width=40)
        self.pdf_page_label.pack(side="left", padx=5)

        self.pdf_next_button = ctk.CTkButton(
            pdf_top_bar, text=">", width=30, state="disabled",
            command=self.show_next_page
        )
        self.pdf_next_button.pack(side="left", padx=5)

        # PDFプレビュー本体 (ラベル)
        self.pdf_preview_label = ctk.CTkLabel(
            self.pdf_preview_container,
            text="プレビューを読込中...",
            fg_color="gray20",
            anchor="center",
            text_color="gray70"
        )
        self.pdf_preview_label.grid(
            row=1, column=0, padx=5, pady=(0, 5), sticky="nsew"
        )

        # --- [Col 1] 情報/メモ/引用元コンテナ ---
        self.info_container = ctk.CTkScrollableFrame(
            self, label_text="ノート詳細"
        )
        self.info_container.grid_columnconfigure(1, weight=1)

        # メモと引用元のために行の重みを設定
        self.info_container.grid_rowconfigure(4, weight=1)  # メモ
        self.info_container.grid_rowconfigure(6, weight=1)  # 引用元

        # --- ウィジェットの作成  ---
        # 1. タイトル
        ctk.CTkLabel(
            self.info_container, text="タイトル:", anchor="w"
            ).grid(row=0, column=0, padx=10, pady=5, sticky="w")
        ctk.CTkLabel(
            self.info_container, text=title, wraplength=300,
            justify="left", anchor="w"
            ).grid(row=0, column=1, padx=10, pady=5, sticky="ew")

        # 2. キー
        ctk.CTkLabel(
            self.info_container, text="キー:", anchor="w"
            ).grid(row=1, column=0, padx=10, pady=5, sticky="w")
        ctk.CTkLabel(
            self.info_container, text=self.note_data.get('key', ''), anchor="w"
            ).grid(row=1, column=1, padx=10, pady=5, sticky="ew")

        # 3. Index Key
        ctk.CTkLabel(
            self.info_container, text="インデックス キー:", anchor="w"
            ).grid(row=2, column=0, padx=10, pady=5, sticky="w")
        ctk.CTkLabel(
            self.info_container,
            text=self.note_data.get('commonplace_key', ''), anchor="w"
            ).grid(row=2, column=1, padx=10, pady=5, sticky="ew")

        # 4. タグ
        ctk.CTkLabel(
            self.info_container, text="タグ:", anchor="w"
            ).grid(row=3, column=0, padx=10, pady=5, sticky="w")
        tags_str = str(self.note_data.get('tags', ''))
        tags_list = [tag for tag in tags_str.split(';') if tag]
        tags_display = ", ".join(tags_list)
        ctk.CTkLabel(
            self.info_container, text=tags_display, wraplength=300,
            justify="left", anchor="w"
            ).grid(row=3, column=1, padx=10, pady=5, sticky="ew")

        # 5. メモ (ラベル)
        ctk.CTkLabel(
            self.info_container, text="メモ:", anchor="w"
            ).grid(row=4, column=0, padx=10, pady=5, sticky="nw")

        # 6. メモ (フレーム)
        self.memo_display_frame = ctk.CTkScrollableFrame(self.info_container)
        self.memo_display_frame.grid(
            row=4, column=1, padx=10, pady=5, sticky="nsew"
            )

        # 7. 引用元 (ラベル)
        ctk.CTkLabel(
            self.info_container, text="引用元:", anchor="w"
            ).grid(row=5, column=0, padx=10, pady=5, sticky="nw")

        # 8. 引用元 (フレーム)
        self.references_display_frame = ctk.CTkScrollableFrame(
            self.info_container, label_text="このノートを引用"
            )
        self.references_display_frame.grid(
            row=6, column=1, padx=10, pady=5, sticky="nsew"
            )

        # --- [Row 1] ボタンエリア ---
        self.button_frame = ctk.CTkFrame(self, fg_color="transparent")

        # PDFを開くボタン
        pdf_button = ctk.CTkButton(
            self.button_frame, text="PDFを開く", command=self.open_pdf_action
        )
        pdf_button.pack(side="left", padx=10)

        # 編集ボタン
        edit_button = ctk.CTkButton(
            self.button_frame,
            text="編集する",
            command=self.edit_note_action,
            fg_color="#585a9c",
            hover_color="#494B83"
        )
        edit_button.pack(side="left", padx=10)

        self.toggle_view_button = ctk.CTkButton(
            self.button_frame,
            text="表示切替",  # テキストは update_layout で設定
            command=self.toggle_view_mode
        )
        self.toggle_view_button.pack(side="left", padx=10)

        # --- PDFプレビューの読み込み処理 ---
        (
            self.current_pdf_doc,
            self.current_pdf_doc_start_index,
            self.current_pdf_page_count
        ) = get_pdf_document_for_note(
            self.note_data,
            self.parent_app.loaded_db_path,
            self.parent_app.pdf_root_folder
        )

        # 1ページ目 (相対インデックス 0) を表示
        self.current_pdf_page_index = 0

        # レイアウトを適用
        self.update_layout()

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

    def toggle_view_mode(self):
        """ウィンドウの表示モードを「全情報」と「PDFのみ」で切り替える"""
        # 状態を反転
        self.is_compact_view = not self.is_compact_view

        # レイアウトを再構築
        self.update_layout()

    def update_layout(self):
        """
        self.is_compact_view の状態に基づき、ウィンドウ全体のレイアウトを再構築する。
        """

        # 1. すべてのメインコンテナを非表示
        self.pdf_preview_container.grid_forget()
        self.info_container.grid_forget()
        self.button_frame.grid_forget()

        # 2. グリッド構成をリセット
        self.grid_columnconfigure(0, weight=0)
        self.grid_columnconfigure(1, weight=0)
        self.grid_rowconfigure(0, weight=0)
        self.grid_rowconfigure(1, weight=0)
        self.grid_rowconfigure(2, weight=0)

        # 3. メモ欄と引用元欄の中身を再構築 (幅が変わるため)
        self._build_references_display()

        if self.is_compact_view:
            # --- 垂直レイアウト (情報カード) ---
            self.geometry(self.compact_geometry)
            # self.transient(self.parent_app)

            # グリッド (1列 x 3行)
            self.grid_columnconfigure(0, weight=1)
            self.grid_rowconfigure(0, weight=4)  # PDF
            self.grid_rowconfigure(1, weight=1)  # Info
            self.grid_rowconfigure(2, weight=0)  # Buttons

            # ウィジェット配置
            self.pdf_preview_container.grid(
                row=0, column=0, padx=10, pady=(10, 5), sticky="nsew"
            )
            self.info_container.grid(
                row=1, column=0, padx=10, pady=5, sticky="nsew"
            )
            self.button_frame.grid(
                row=2, column=0, padx=10, pady=10, sticky="ew"
            )

            self.toggle_view_button.configure(text="拡大表示")

            # メモ欄の幅を 400 に設定して再構築
            self._build_memo_display(frame_width=400)

            # PDFプレビューの幅を 250 に設定して更新
            self.update_pdf_preview_image(max_width_override=400)

        else:
            # --- 水平レイアウト (詳細プレビュー) ---
            self.geometry(self.full_geometry)

            # グリッド (2列 x 2行)
            self.grid_columnconfigure(0, weight=1) # PDF
            self.grid_columnconfigure(1, weight=2) # Info
            self.grid_rowconfigure(0, weight=1) # Main
            self.grid_rowconfigure(1, weight=0) # Buttons

            # ウィジェット配置
            self.pdf_preview_container.grid(
                row=0, column=0, padx=(10, 5), pady=10, sticky="nsew"
            )
            self.info_container.grid(
                row=0, column=1, padx=(5, 10), pady=10, sticky="nsew"
            )
            self.button_frame.grid(
                row=1, column=0, columnspan=2, padx=10, pady=10
            )

            self.toggle_view_button.configure(text="縮小表示")

            # メモ欄の幅を 800 に設定して再構築
            self._build_memo_display(frame_width=800)

            # PDFプレビューの幅を 400 に設定して更新
            self.update_pdf_preview_image(max_width_override=400)

    def on_close(self):
        """ ウィンドウを閉じる際の処理 """
        # PyMuPDF (fitz) ドキュメントを閉じる
        if self.current_pdf_doc:
            try:
                self.current_pdf_doc.close()
                self.current_pdf_doc = None
            except Exception as e:
                logger.error(f"プレビューのPDFドキュメント解放エラー: {e}")

        self.destroy()

    def open_pdf_action(self):
        """「PDFを開く」ボタンが押されたときの処理。"""
        self.parent_app.open_pdf(self.note_data)
        self.on_close()

    def edit_note_action(self):
        """
        編集ボタンが押されたときの処理。
        """
        self.on_close()
        self.parent_app.open_edit_dialog(self.note_data)

    def _build_memo_display(self, frame_width=400):
        """
        メモ欄にクリック可能なリンク付きラベルを生成する。
        """
        memo_text = str(self.note_data.get('memo', ''))

        build_memo_display(
            self.memo_display_frame,
            memo_text,
            self.parent_app.df,
            self.parent_app.open_preview_window,
            frame_width
        )

    def _build_references_display(self):
        """
        プレビューウィンドウの引用元欄を構築する。
        """
        current_key = self.note_data.get('key', '')

        backlinks_df = pd.DataFrame()  # 空で初期化
        # メインアプリ (parent_app) からDB接続を取得
        db_conn = self.parent_app.db_conn

        if db_conn and current_key:
            try:
                # メインウィンドウの show_details と同じ LIKE 検索を実行
                # リンクテーブルを検索
                sql = "SELECT source_key FROM note_links WHERE target_key = ?"
                cursor = db_conn.cursor()
                cursor.execute(sql, (current_key,))

                matching_keys = {row[0] for row in cursor.fetchall()}

                if matching_keys:
                    backlinks_df = (
                        self.parent_app.df[
                            self.parent_app.df['key'].isin(matching_keys)
                        ]
                    )
            except Exception as e:
                logger.error(f"[Preview] 引用元のDB検索エラー: {e}", exc_info=True)
                pass

        # utilsの関数を使って引用元UIを構築
        build_references_display(
            self.references_display_frame,
            backlinks_df,
            self.parent_app.open_preview_window,
            self.parent_app.key_icons,
            self.parent_app.key_colors
        )

    def update_pdf_preview_image(self, max_width_override=None):
        """
        現在のPDFドキュメントとページインデックスに基づき、
        プレビュー画像とページめくりUIを更新する。
        """
        if self.current_pdf_doc is None:
            self.preview_image_object = None
            self.pdf_preview_label.configure(
                image=None, text="プレビューの読み込みに失敗",
                fg_color="gray20", text_color="#D9534F"
            )
            self.pdf_page_label.configure(text="(-/-)")
            self.pdf_prev_button.configure(state="disabled")
            self.pdf_next_button.configure(state="disabled")
            return

        # 1. 表示すべき絶対ページインデックスを計算
        absolute_page_index = (
            self.current_pdf_doc_start_index + self.current_pdf_page_index
        )

        # 2. 画像取得
        max_preview_width = max_width_override if max_width_override else 400

        pil_image = get_pdf_page_image_from_doc(
            self.current_pdf_doc,
            absolute_page_index,
            max_width=max_preview_width
        )

        # 3. 画像をUIに設定
        if pil_image:
            self.preview_image_object = ctk.CTkImage(
                light_image=pil_image,
                dark_image=pil_image,
                size=(pil_image.width, pil_image.height)
            )
            self.pdf_preview_label.configure(
                image=self.preview_image_object,
                text="",
                fg_color="transparent"
            )
        else:
            self.preview_image_object = None
            self.pdf_preview_label.configure(
                image=None,
                text=f"ページ {absolute_page_index + 1} の描画に失敗",
                fg_color="gray20",
                text_color="#D9534F"
            )

        # 4. ページめくりUIを更新 (1-indexed)
        self.pdf_page_label.configure(
            text=(
                f"{self.current_pdf_page_index + 1} / "
                f"{self.current_pdf_page_count}"
            )
        )
        # 「前へ」ボタン
        if self.current_pdf_page_index > 0:
            self.pdf_prev_button.configure(state="normal")
        else:
            self.pdf_prev_button.configure(state="disabled")
        # 「次へ」ボタン
        if self.current_pdf_page_index < (self.current_pdf_page_count - 1):
            self.pdf_next_button.configure(state="normal")
        else:
            self.pdf_next_button.configure(state="disabled")

    def show_prev_page(self):
        """「<」ボタン: 前のページを表示"""
        if self.current_pdf_doc and self.current_pdf_page_index > 0:
            self.current_pdf_page_index -= 1
            self.update_layout()

    def show_next_page(self):
        """「>」ボタン: 次のページを表示"""
        if (
            self.current_pdf_doc and
            self.current_pdf_page_index < (self.current_pdf_page_count - 1)
        ):
            self.current_pdf_page_index += 1
            self.update_layout()

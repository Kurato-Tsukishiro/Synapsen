from tkinter import messagebox
import customtkinter as ctk
import fitz
import logging
import sys
from pathlib import Path

from utils import (
    build_memo_display,
    build_references_display,
    get_pdf_document_for_note,
    get_pdf_page_image_from_doc,
    _extract_links,
)

logger = logging.getLogger(__name__)

# === 2. プロジェクトルートをパスに追加 ===
current_dir = Path(__file__).parent
root_dir = current_dir.parent
if str(root_dir) not in sys.path:
    sys.path.append(str(root_dir))

from theme import SemanticColors as Colors  # noqa: E402


# ==============================================================================
# 簡易プレビューウィンドウ (読み取り専用)
# ==============================================================================
class NotePreviewWindow(ctk.CTkToplevel):
    """
    ノートのメタデータをプレビュー表示するための専用Toplevelウィンドウ。
    メインウィンドウから独立して表示され、ノートの詳細とPDFへのショートカットを提供する。
    """

    def __init__(
        self, parent_app, note_data, default_view_mode="compact", ui_master=None
    ):
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
            ui_master (DigitalCommonplaceBook):
                呼び出し元、ui_master が指定されていればそれを親(master)にする、なければ parent_app
        """
        master_window = ui_master if ui_master else parent_app
        super().__init__(parent_app)
        self.parent_app = parent_app  # メインアプリ本体

        # 辞書化を保証
        if not isinstance(note_data, dict):
            try:
                self.note_data = dict(note_data)
            except Exception:
                self.note_data = {}
        else:
            self.note_data = note_data

        # CTkImageオブジェクトへの参照を保持 (ガベージコレクション対策)
        self.preview_image_object = None

        # ページめくり用の状態変数
        self.current_pdf_doc: fitz.Document | None = None
        self.current_pdf_page_index = 0
        self.current_pdf_page_count = 0
        self.current_pdf_doc_start_index = 0

        # メモ欄から抽出された引用先リンクのセット
        self.forward_links = set()

        # プレビューウィンドウのレイアウト状態とサイズを管理
        self.is_compact_view = default_view_mode == "compact"
        self.compact_geometry = "450x750"  # 垂直レイアウト    (リンク呼び出し)
        self.full_geometry = "1350x750"  # 2カラムレイアウト (本体呼び出し)

        self._custom_icon_path = None
        if hasattr(parent_app, "icon_path") and parent_app.icon_path:
            self._custom_icon_path = str(parent_app.icon_path)

            if self._custom_icon_path:
                try:
                    super().iconbitmap(self._custom_icon_path)
                except Exception as e:
                    logger.error(f"Initial icon set error: {e}")

        title = self.note_data.get("title", "N/A")
        self.title(f"プレビュー: {title}")
        self.transient(master_window)

        self.protocol("WM_DELETE_WINDOW", self.on_close)

        # --- [Col 0] PDFプレビューコンテナ ---
        self.pdf_preview_container = ctk.CTkFrame(
            self, fg_color=Colors.BACKGROUND_PANEL
        )
        self.pdf_preview_container.grid_rowconfigure(1, weight=1)
        self.pdf_preview_container.grid_columnconfigure(0, weight=1)

        # PDFプレビュー用の上部バー (ラベル + ページめくり)
        pdf_top_bar = ctk.CTkFrame(self.pdf_preview_container, fg_color="transparent")
        pdf_top_bar.grid(row=0, column=0, padx=5, pady=5, sticky="ew")
        ctk.CTkLabel(pdf_top_bar, text="PDFプレビュー:", anchor="w").pack(
            side="left", padx=(0, 10)
        )

        # ページめくりボタン (初期状態は無効)
        self.pdf_prev_button = ctk.CTkButton(
            pdf_top_bar,
            text="<",
            width=30,
            state="disabled",
            fg_color=Colors.UI_BASIC,
            hover_color=Colors.adjust_brightness(Colors.UI_BASIC),
            text_color=Colors.adjust_brightness(Colors.UI_BASIC, factor=0.2),
            command=self.show_prev_page,
        )
        self.pdf_prev_button.pack(side="left", padx=5)

        self.pdf_page_label = ctk.CTkLabel(pdf_top_bar, text="(-/-)", width=40)
        self.pdf_page_label.pack(side="left", padx=5)

        self.pdf_next_button = ctk.CTkButton(
            pdf_top_bar,
            text=">",
            width=30,
            state="disabled",
            fg_color=Colors.UI_BASIC,
            hover_color=Colors.adjust_brightness(Colors.UI_BASIC),
            text_color=Colors.adjust_brightness(Colors.UI_BASIC, factor=0.2),
            command=self.show_next_page,
        )
        self.pdf_next_button.pack(side="left", padx=5)

        # PDFプレビュー本体 (ラベル)
        self.pdf_preview_label = ctk.CTkLabel(
            self.pdf_preview_container,
            text="プレビューを読込中...",
            fg_color="gray20",
            anchor="center",
            text_color="gray70",
        )
        self.pdf_preview_label.grid(row=1, column=0, padx=5, pady=(0, 5), sticky="nsew")

        # --- [Col 1] 情報/メモ/引用元コンテナ ---
        self.info_container = ctk.CTkScrollableFrame(
            self,
            label_text="ノート詳細",
            fg_color=Colors.BACKGROUND_PANEL,
            label_fg_color=Colors.adjust_brightness(Colors.BACKGROUND_PANEL, 0.8),
        )
        self.info_container.grid_columnconfigure(1, weight=1)

        # メモと引用元のために行の重みを設定
        self.info_container.grid_rowconfigure(5, weight=1)  # メモ
        self.info_container.grid_rowconfigure(7, weight=1)  # 引用元

        # --- ウィジェットの作成 ---
        # 1. タイトル
        ctk.CTkLabel(self.info_container, text="タイトル:", anchor="w").grid(
            row=0, column=0, padx=10, pady=5, sticky="w"
        )
        ctk.CTkLabel(
            self.info_container, text=title, wraplength=300, justify="left", anchor="w"
        ).grid(row=0, column=1, padx=10, pady=5, sticky="ew")

        # 2. キー
        ctk.CTkLabel(self.info_container, text="キー:", anchor="w").grid(
            row=1, column=0, padx=10, pady=5, sticky="w"
        )
        ctk.CTkLabel(
            self.info_container, text=self.note_data.get("key", ""), anchor="w"
        ).grid(row=1, column=1, padx=10, pady=5, sticky="ew")

        # 3. Index Key
        ctk.CTkLabel(self.info_container, text="インデックス キー:", anchor="w").grid(
            row=2, column=0, padx=10, pady=5, sticky="w"
        )
        ctk.CTkLabel(
            self.info_container,
            text=self.note_data.get("commonplace_key", ""),
            anchor="w",
        ).grid(row=2, column=1, padx=10, pady=5, sticky="ew")

        # 4. 概要
        ctk.CTkLabel(self.info_container, text="概要:", anchor="w").grid(
            row=3, column=0, padx=10, pady=5, sticky="w"
        )
        ctk.CTkLabel(
            self.info_container,
            text=self.note_data.get("summary", ""),
            wraplength=300,
            justify="left",
            anchor="w",
        ).grid(row=3, column=1, padx=10, pady=5, sticky="ew")

        # 5. タグ
        ctk.CTkLabel(self.info_container, text="タグ:", anchor="w").grid(
            row=4, column=0, padx=10, pady=5, sticky="w"
        )
        tags_str = str(self.note_data.get("tags", ""))
        tags_list = [tag for tag in tags_str.split(";") if tag]
        tags_display = ", ".join(tags_list)
        ctk.CTkLabel(
            self.info_container,
            text=tags_display,
            wraplength=300,
            justify="left",
            anchor="w",
        ).grid(row=4, column=1, padx=10, pady=5, sticky="ew")

        # 6. メモ
        ctk.CTkLabel(self.info_container, text="メモ:", anchor="w").grid(
            row=5, column=0, padx=10, pady=5, sticky="nw"
        )

        # メモと引用元のエリアのラベル色
        label_fg_color = Colors.adjust_brightness(Colors.BACKGROUND_HOLLOW, 0.85)

        # 7. メモ (フレーム)
        self.memo_display_frame = ctk.CTkScrollableFrame(
            self.info_container,
            fg_color=Colors.BACKGROUND_HOLLOW,
            label_fg_color=label_fg_color,
        )
        self.memo_display_frame.grid(row=5, column=1, padx=10, pady=5, sticky="nsew")

        # 8. 引用元
        ctk.CTkLabel(self.info_container, text="引用元:", anchor="w").grid(
            row=6, column=0, padx=10, pady=5, sticky="nw"
        )

        # 9. 引用元 (フレーム)
        self.references_display_frame = ctk.CTkScrollableFrame(
            self.info_container,
            label_text="このノートを引用",
            fg_color=Colors.BACKGROUND_HOLLOW,
            label_fg_color=label_fg_color,
        )
        self.references_display_frame.grid(
            row=7, column=1, padx=10, pady=5, sticky="nsew"
        )

        # --- [Row 1] ボタンエリア ---
        self.button_frame = ctk.CTkFrame(self, fg_color="transparent")

        # 10. PDFを開くボタン
        self.pdf_button = ctk.CTkButton(
            self.button_frame, text="PDFを開く",
            fg_color=Colors.UI_PREVIEW,
            hover_color=Colors.adjust_brightness(Colors.UI_PREVIEW),
            command=self.open_pdf_action, width=50
        )
        self.pdf_button.pack(side="left", padx=5)

        # 11. ウィンドウサイズ切り替えボタン
        self.toggle_view_button = ctk.CTkButton(
            self.button_frame,
            text="表示切替",  # テキストは update_layout で設定
            fg_color=Colors.UI_BASIC,
            hover_color=Colors.adjust_brightness(Colors.UI_BASIC),
            text_color=Colors.adjust_brightness(Colors.UI_BASIC, factor=0.2),
            command=self.toggle_view_mode,
        )
        self.toggle_view_button.pack(side="left", padx=5)

        # 12. 関連グラフボタン
        self.graph_button = ctk.CTkButton(
            self.button_frame,
            text="関連グラフ",
            command=self.show_local_graph_action,
            fg_color=Colors.UI_SECONDARY,
            hover_color=Colors.adjust_brightness(Colors.UI_SECONDARY),
        )
        self.graph_button.pack(side="left", padx=5)

        # 13. 本体Key/引用先コピーボタン
        self.copy_menu_var = ctk.StringVar(value="コピー...")
        self.copy_menu = ctk.CTkOptionMenu(
            self.button_frame,
            variable=self.copy_menu_var,
            values=["本体のKeyをコピー", "引用先をコピー"],
            command=self.handle_copy_menu,
            fg_color=Colors.UI_LINK,  # ボタン色
            button_color=Colors.adjust_brightness(  # ドロップダウン矢印の色
                Colors.UI_LINK
            ),
            button_hover_color=Colors.adjust_brightness(  # ドロップダウン矢印のホバー色
                Colors.UI_LINK, 0.6
            ),
        )
        self.copy_menu.pack(side="left", padx=5)

        # 14. 編集ボタン
        self.edit_button = ctk.CTkButton(
            self.button_frame,
            text="編集する",
            command=self.edit_note_action,
            fg_color=Colors.UI_EDIT,
            hover_color=Colors.adjust_brightness(Colors.UI_EDIT),
        )
        self.edit_button.pack(side="left", padx=5)

        # 15. 選択ノートへリンク
        self.link_to_selected_button = ctk.CTkButton(
            self.button_frame,
            text="選択へリンク",
            command=self.link_to_selected_action,
            fg_color=Colors.UI_EDIT,
            hover_color=Colors.adjust_brightness(Colors.UI_EDIT),
        )
        self.link_to_selected_button.pack(side="left", padx=5)

        # --- PDFプレビューの読み込み処理 ---
        (
            self.current_pdf_doc,
            self.current_pdf_doc_start_index,
            self.current_pdf_page_count,
        ) = get_pdf_document_for_note(
            self.note_data,
            self.parent_app.loaded_db_path,
            self.parent_app.pdf_root_folder,
            pdf_archive_folder=self.parent_app.pdf_archive_folder,
        )

        # 1ページ目 (相対インデックス 0) を表示
        self.current_pdf_page_index = 0

        # レイアウトを適用
        self.update_layout()

    def show_local_graph_action(self):
        """「関連グラフ」ボタンが押されたときの処理"""
        key_to_show = self.note_data.get("key")
        if key_to_show:
            self.parent_app.show_local_graph(key_to_show)
        else:
            messagebox.showwarning(
                "キー不明", "ノートKeyが不明なためグラフを表示できません。", parent=self
            )

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
        """ウィンドウ全体のレイアウトを再構築する"""
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

        # メモ欄と引用元欄の再構築 (df参照を排除したメソッドを使用)
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
            self.info_container.grid(row=1, column=0, padx=10, pady=5, sticky="nsew")
            self.button_frame.grid(row=2, column=0, padx=10, pady=10, sticky="ew")

            self.toggle_view_button.configure(text="拡大表示")

            self.pdf_button.configure(text="PDF", width=60)
            self.edit_button.configure(text="編集", width=60)
            self.toggle_view_button.configure(text="拡大", width=60)
            self.graph_button.configure(text="グラフ", width=60)
            self.copy_menu_var.set("コピー")
            self.copy_menu.configure(width=60)
            self.link_to_selected_button.configure(text="リンク", width=60)

            self._build_memo_display(frame_width=400)
            self.update_pdf_preview_image(max_width_override=400)
        else:
            # --- 水平レイアウト (詳細プレビュー) ---
            self.geometry(self.full_geometry)

            # グリッド (2列 x 2行)
            self.grid_columnconfigure(0, weight=1)  # PDF
            self.grid_columnconfigure(1, weight=2)  # Info
            self.grid_rowconfigure(0, weight=1)  # Main
            self.grid_rowconfigure(1, weight=0)  # Buttons

            # ウィジェット配置
            self.pdf_preview_container.grid(
                row=0, column=0, padx=(10, 5), pady=10, sticky="nsew"
            )
            self.info_container.grid(
                row=0, column=1, padx=(5, 10), pady=10, sticky="nsew"
            )
            self.button_frame.grid(row=1, column=0, columnspan=2, padx=10, pady=10)

            self.toggle_view_button.configure(text="縮小表示")

            # ★ ボタンのテキストを元に戻す
            self.pdf_button.configure(text="PDFを開く", width=140)
            self.edit_button.configure(text="編集する", width=140)
            self.toggle_view_button.configure(text="縮小表示", width=140)
            self.graph_button.configure(text="関連グラフ", width=140)
            self.copy_menu_var.set("コピー...")
            self.copy_menu.configure(width=140)
            self.link_to_selected_button.configure(
                text="選択ノートへリンクを付与", width=140
            )

            self._build_memo_display(frame_width=800)
            self.update_pdf_preview_image(max_width_override=400)

    def on_close(self):
        """ウィンドウを閉じる際の処理"""
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

        # 現在が「拡大表示」モード (is_compact_view == False) だったら
        if not self.is_compact_view:
            # 状態を「縮小表示」モード (True) に設定
            self.is_compact_view = True
            self.update_layout()

    def edit_note_action(self):
        """
        編集ボタンが押されたときの処理。
        """
        self.on_close()
        self.parent_app.open_edit_dialog(self.note_data)

    def handle_copy_menu(self, choice: str):
        """OptionMenuで項目が選択されたときの処理"""
        if choice == "本体のKeyをコピー":
            self.copy_own_key_action()
        elif choice == "引用先をコピー":
            self.copy_forward_links_action()

        # 選択後にメニューの表示を元に戻す
        if self.is_compact_view:
            self.copy_menu_var.set("コピー")
        else:
            self.copy_menu_var.set("コピー...")

    def copy_forward_links_action(self):
        """「引用先コピー」ボタンの処理 (DB対応版)"""
        if not self.forward_links:
            messagebox.showinfo(
                "情報",
                "このノートのメモ欄には引用先リンク ([[...]]) がありません。",
                parent=self,
            )
            return

        conn = self.parent_app.db_conn
        if not conn:
            return

        try:
            target_keys = list(self.forward_links)
            if not target_keys:
                return

            placeholders = ",".join("?" * len(target_keys))
            sql = (
                "SELECT key, title FROM notes "
                f"WHERE key IN ({placeholders}) ORDER BY key"
            )

            cursor = conn.cursor()
            cursor.execute(sql, target_keys)
            rows = cursor.fetchall()

            link_texts = []
            found_keys = set()
            for r in rows:
                key, title = r[0], r[1]
                link_texts.append(f"[[{key}: {title}]]")
                found_keys.add(key)

            # 見つからなかったキーも含める
            for k in target_keys:
                if k not in found_keys:
                    link_texts.append(f"[[{k}]]")

            clipboard_text = "\n".join(link_texts)
            self.clipboard_clear()
            self.clipboard_append(clipboard_text)
            self.update()

            messagebox.showinfo(
                "コピー完了",
                f"{len(link_texts)}件の引用先リンクをクリップボードにコピーしました。",
                parent=self,
            )
        except Exception as e:
            logger.error(f"引用先コピーエラー: {e}")
            messagebox.showerror("エラー", f"コピーに失敗しました: {e}")

    def copy_own_key_action(self):
        """「Keyコピー」ボタンが押されたときの処理"""
        key_to_copy = self.note_data.get("key")
        title_to_copy = self.note_data.get("title")  # ★ Titleも取得

        if not key_to_copy:
            messagebox.showwarning(
                "キー不明", "このノートのKeyが不明です。", parent=self
            )
            return

        # リンク形式の文字列を生成
        clipboard_text = f"[[{key_to_copy}: {title_to_copy}]]"

        # クリップボードへコピー
        self.clipboard_clear()
        self.clipboard_append(clipboard_text)
        self.update()

        messagebox.showinfo(
            "コピー完了", f"リンク形式でコピーしました:\n{clipboard_text}", parent=self
        )

    def link_to_selected_action(self):
        """「選択ノートへリンク」ボタンが押されたときの処理"""
        key_to_link = self.note_data.get("key")
        title_to_link = self.note_data.get("title", "")

        if not key_to_link:
            messagebox.showwarning(
                "キー不明", "Keyが不明なためリンクできません。", parent=self
            )
            return

        self.parent_app.append_link_to_selected_notes(key_to_link, title_to_link)

    def _build_memo_display(self, frame_width=400):
        """
        メモ欄にクリック可能なリンク付きラベルを生成し
        引用先キー(self.forward_links)を抽出する。
        """
        memo_text = str(self.note_data.get("memo", ""))

        # メモを解析して引用先キーのセットを更新
        # メモはすでに取得済みの為、わざわざ"note_links" テーブルから取得しない
        self.forward_links = _extract_links(memo_text)

        build_memo_display(
            self.memo_display_frame,
            memo_text,
            self.parent_app.db_conn,
            lambda key: self.parent_app.open_preview_window(
                key, default_view_mode="compact", ui_master=self
            ),
            frame_width,
        )

    def _build_references_display(self):
        """
        プレビューウィンドウの引用元欄を構築する。
        """
        current_key = self.note_data.get("key", "")
        backlinks_data = []
        db_conn = self.parent_app.db_conn

        if db_conn and current_key:
            try:
                sql = """
                    SELECT n.key, n.title, n.date, n.commonplace_key
                    FROM notes n
                    JOIN note_links l ON n.key = l.source_key
                    WHERE l.target_key = ?
                    ORDER BY n.date DESC, n.time DESC
                """
                cursor = db_conn.cursor()
                cursor.execute(sql, (current_key,))
                rows = cursor.fetchall()

                for r in rows:
                    backlinks_data.append(
                        {
                            "key": r[0],
                            "title": r[1],
                            "date": r[2],
                            "commonplace_key": r[3],
                        }
                    )
            except Exception as e:
                logger.error(f"[Preview] 引用元のDB検索エラー: {e}", exc_info=True)

        build_references_display(
            self.references_display_frame,
            backlinks_data,
            lambda key: self.parent_app.open_preview_window(
                key, default_view_mode="compact", ui_master=self
            ),
            self.parent_app.key_icons,
            self.parent_app.key_colors,
        )

    def update_pdf_preview_image(self, max_width_override=None):
        """
        現在のPDFドキュメントとページインデックスに基づき、
        プレビュー画像とページめくりUIを更新する。
        """
        if self.current_pdf_doc is None:
            self.preview_image_object = None
            self.pdf_preview_label.configure(
                image=None,
                text="プレビューの読み込みに失敗",
                fg_color="gray20",
                text_color=Colors.LABEL_DENGER,
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
            self.current_pdf_doc, absolute_page_index, max_width=max_preview_width
        )

        # 3. 画像をUIに設定
        if pil_image:
            self.preview_image_object = ctk.CTkImage(
                light_image=pil_image,
                dark_image=pil_image,
                size=(pil_image.width, pil_image.height),
            )
            self.pdf_preview_label.configure(
                image=self.preview_image_object, text="", fg_color="transparent"
            )
        else:
            self.preview_image_object = None
            self.pdf_preview_label.configure(
                image=None,
                text=f"ページ {absolute_page_index + 1} の描画に失敗",
                fg_color="gray20",
                text_color=Colors.LABEL_DENGER,
            )

        # 4. ページめくりUIを更新 (1-indexed)
        self.pdf_page_label.configure(
            text=(
                f"{self.current_pdf_page_index + 1} / " f"{self.current_pdf_page_count}"
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
        if self.current_pdf_doc and self.current_pdf_page_index < (
            self.current_pdf_page_count - 1
        ):
            self.current_pdf_page_index += 1
            self.update_layout()

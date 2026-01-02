import sys
import customtkinter as ctk
from tkinter import filedialog, messagebox
from pathlib import Path
import datetime
import re
import tkinterdnd2
from PIL import ImageGrab, Image
import tempfile
import shutil
from pypdf import PdfReader, PdfWriter
import fitz  # PDFプレビュー用
import io

# 追加した画像エディタをインポート
from image_editor import PerspectiveCropEditor

from pdf_utils import (
    # D&D/画像クリップ用のメタデータ挿入関数
    add_metadata_to_clip,
    hex_to_rgb_tuple,
    # D&Dパイプラインで個別に実行するためインポート
    convert_image_to_pdf,
    convert_pil_image_to_pdf,
    convert_document_to_pdf,
    high_fidelity_flatten,
    normalize_pdf_to_papersize,
    embed_ocr_text_in_pdf,
)

import logging

logger = logging.getLogger(__name__)

# === 2. プロジェクトルートをパスに追加 ===
current_dir = Path(__file__).parent
root_dir = current_dir.parent
if str(root_dir) not in sys.path:
    sys.path.append(str(root_dir))

from theme import SemanticColors as Colors  # noqa: E402

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


class DragAndDropWindow(ctk.CTkToplevel, tkinterdnd2.TkinterDnD.DnDWrapper):
    """
    ドラッグ＆ドロップとペーストによる正規化（ファイル名編集機能付き）
    を行うためのToplevelウィンドウ（モーダル）。

    Attributes:
        parent_app (ctk.CTk): 親である Synapsen_Normalisierer のインスタンス。
        staged_items (list[dict]):
            処理対象としてステージングされたアイテムのリスト。
            各辞書は "data" (Path or PIL.Image), "base_name_var" (ctk.StringVar),
            "original_name" (str) をキーとして持つ。
    """

    def __init__(self, parent_app):
        """
        DragAndDropWindowを初期化します。

        Args:
            parent_app (ctk.CTk): 親ウィンドウ (Synapsen_Normalisierer)。
        """
        super().__init__(parent_app)

        # tkinterdnd2 の初期化
        try:
            self.TkdndVersion = tkinterdnd2.TkinterDnD._require(self)
        except Exception as e:
            messagebox.showerror(
                "DND初期化エラー",
                f"tkinterdnd2 の初期化に失敗しました: {e}",
                parent=self,
            )
            self.destroy()
            return

        self.parent_app = parent_app  # メインアプリ(Synapsen_Normalisierer)
        self.staged_items = []

        # ポップアップウィンドウの参照保持用
        self.preview_popup = None
        self.preview_timer = None  # プレビュー遅延表示用のタイマー
        self.preview_label = None  # ラベルウィジェットの参照を保持
        self.preview_image_cache = {}  # アイテムごとのサムネイルキャッシュ
        self.hovering_item = None  # 現在ホバー中のアイテムを管理する変数
        self.drag_source_item = None

        self.title("D&D/ペーストで正規化")
        self.geometry("1000x800")
        self.configure(
            fg_color=(Colors.BACKGROUND_HOLLOW, Colors.BACKGROUND_DARK_HOLLOW)
        )

        self._custom_icon_path = None
        if hasattr(parent_app, "icon_path") and parent_app.icon_path:
            self._custom_icon_path = str(parent_app.icon_path)
            # ウィンドウ生成直後のリセットを防ぐため、少し遅延させて適用
            self.after(200, lambda: self.iconbitmap(default=self._custom_icon_path))

        self.protocol("WM_DELETE_WINDOW", self.on_close)
        self.transient(parent_app)
        self.grab_set()

        # --- [UI定義] ---
        self.grid_rowconfigure(1, weight=1)  # 2カラムのメインフレーム領域
        self.grid_columnconfigure(0, weight=1)

        # 1. ドラッグ＆ドロップ ゾーン
        self.drop_target_frame = ctk.CTkFrame(
            self,
            height=100,
            fg_color=Colors.blend_colors("#000000", Colors.BACKGROUND_PANEL, 0.75),
        )
        self.drop_target_frame.grid(row=0, column=0, pady=10, padx=10, sticky="ew")

        self.drop_target_label = ctk.CTkLabel(
            self.drop_target_frame,
            text=(
                "ここにファイル (PDF, JPG, PNG, MD) を\nドラッグ＆ドロップしてください\n\n"
                + "(または Ctrl+V でスクリーンショットを貼付)"
            ),
            text_color="gray70",
        )
        self.drop_target_label.place(relx=0.5, rely=0.5, anchor="center")

        # D&Dイベントのバインド
        self.drop_target_frame.drop_target_register(tkinterdnd2.DND_FILES)
        self.drop_target_frame.dnd_bind("<<Drop>>", self.handle_drop)
        # ペーストイベントのバインド
        self.bind_all("<Control-v>", self.handle_paste)

        main_frame = ctk.CTkFrame(self, fg_color="transparent")
        main_frame.grid(row=1, column=0, pady=0, padx=10, sticky="nsew")
        main_frame.grid_rowconfigure(0, weight=1)
        main_frame.grid_columnconfigure(0, weight=1)  # 左カラム (メタデータ)
        main_frame.grid_columnconfigure(1, weight=1)  # 右カラム (リスト)

        # 2. メタデータ入力フレーム (左カラム)
        meta_frame = ctk.CTkFrame(main_frame, fg_color=Colors.BACKGROUND_PANEL)
        meta_frame.grid(row=0, column=0, pady=0, padx=(0, 5), sticky="nsew")

        # IndexKey
        ctk.CTkLabel(
            meta_frame,
            text="IndexKey (PDF 1ページ目に埋込):",
            anchor="w",
        ).pack(pady=(10, 0), padx=10, fill="x")
        key_options = self.parent_app.config_data.get("commonplace_keys_options", [])
        self.index_key_combo = ctk.CTkComboBox(
            meta_frame,
            values=["（未選択）"] + key_options,
            fg_color=(Colors.BACKGROUND_HOLLOW, Colors.BACKGROUND_DARK_HOLLOW),
            button_color=Colors.adjust_brightness(Colors.BACKGROUND_PANEL),
            button_hover_color=Colors.adjust_brightness(Colors.BACKGROUND_PANEL, 0.6),
            dropdown_fg_color=Colors.BACKGROUND_PANEL,
            dropdown_hover_color=Colors.adjust_brightness(Colors.BACKGROUND_PANEL),
        )
        self.index_key_combo.set("（未選択）")
        self.index_key_combo.pack(pady=5, padx=10, fill="x")

        # コメント
        ctk.CTkLabel(
            meta_frame, text="コメント (PDF 最終ページに埋込):", anchor="w"
        ).pack(pady=(5, 0), padx=10, fill="x")
        self.comment_textbox = ctk.CTkTextbox(
            meta_frame,
            height=80,
            fg_color=(Colors.BACKGROUND_HOLLOW, Colors.BACKGROUND_DARK_HOLLOW),
        )
        self.comment_textbox.pack(pady=5, padx=10, fill="both", expand=True)

        # 書誌情報
        ctk.CTkLabel(meta_frame, text="書誌情報 (任意入力)", anchor="w").pack(
            pady=(10, 0), padx=10, fill="x"
        )
        sist_frame = ctk.CTkFrame(
            meta_frame, fg_color=Colors.adjust_brightness(Colors.BACKGROUND_PANEL, 0.9)
        )
        sist_frame.pack(pady=(0, 5), padx=5, fill="x")
        sist_frame.grid_columnconfigure(1, weight=1)

        ctk.CTkLabel(sist_frame, text="著者名:").grid(
            row=0, column=0, padx=5, pady=5, sticky="w"
        )
        self.sist_author_entry = ctk.CTkEntry(
            sist_frame,
            placeholder_text="（任意）",
            fg_color=(Colors.BACKGROUND_HOLLOW, Colors.BACKGROUND_DARK_HOLLOW),
        )
        self.sist_author_entry.grid(row=0, column=1, padx=5, pady=5, sticky="ew")

        ctk.CTkLabel(sist_frame, text="ページ名:").grid(
            row=1, column=0, padx=5, pady=5, sticky="w"
        )
        self.sist_title_entry = ctk.CTkEntry(
            sist_frame,
            placeholder_text="（任意）",
            fg_color=(Colors.BACKGROUND_HOLLOW, Colors.BACKGROUND_DARK_HOLLOW),
        )
        self.sist_title_entry.grid(row=1, column=1, padx=5, pady=5, sticky="ew")

        ctk.CTkLabel(sist_frame, text="サイト名/出典:").grid(
            row=2, column=0, padx=5, pady=5, sticky="w"
        )
        self.sist_site_entry = ctk.CTkEntry(
            sist_frame,
            placeholder_text="（任意）",
            fg_color=(Colors.BACKGROUND_HOLLOW, Colors.BACKGROUND_DARK_HOLLOW),
        )
        self.sist_site_entry.grid(row=2, column=1, padx=5, pady=5, sticky="ew")

        ctk.CTkLabel(sist_frame, text="更新日:").grid(
            row=3, column=0, padx=5, pady=5, sticky="w"
        )
        self.sist_date_entry = ctk.CTkEntry(
            sist_frame,
            placeholder_text="（任意, YYYY-MM-DD）",
            fg_color=(Colors.BACKGROUND_HOLLOW, Colors.BACKGROUND_DARK_HOLLOW),
        )
        self.sist_date_entry.grid(row=3, column=1, padx=5, pady=5, sticky="ew")

        self.is_webclip_var = ctk.BooleanVar(value=False)
        self.webclip_checkbox = ctk.CTkCheckBox(
            meta_frame,
            text="WebClipとして扱う (SType_WebClip を付与)",
            variable=self.is_webclip_var,
            onvalue=True,
            offvalue=False,
            fg_color=Colors.UI_BASIC,
            hover_color=Colors.adjust_brightness(Colors.UI_BASIC),
            checkmark_color=Colors.adjust_brightness(Colors.UI_BASIC, 1.8),
        )
        self.webclip_checkbox.pack(pady=(5, 0), padx=15, anchor="w")

        # 引用Key
        ctk.CTkLabel(
            meta_frame, text="引用元Key (カンマ区切り または 改行区切り):", anchor="w"
        ).pack(pady=(10, 0), padx=10, fill="x")
        self.cited_keys_entry = ctk.CTkTextbox(
            meta_frame,
            height=40,
            fg_color=(Colors.BACKGROUND_HOLLOW, Colors.BACKGROUND_DARK_HOLLOW),
        )
        self.cited_keys_entry.pack(pady=5, padx=10, fill="both", expand=True)

        # 統合チェックボックス
        self.merge_files_checkbox = ctk.CTkCheckBox(
            meta_frame,
            text="処理対象ファイルを1つのノート（PDF）に統合する",
            command=self.on_merge_toggle,
            fg_color=Colors.UI_BASIC,
            hover_color=Colors.adjust_brightness(Colors.UI_BASIC),
            checkmark_color=Colors.adjust_brightness(Colors.UI_BASIC, 1.8),
        )
        self.merge_files_checkbox.pack(pady=10, padx=10, anchor="w")

        # 処理対象リスト (スクロールフレーム)
        self.staged_list_frame = ctk.CTkScrollableFrame(
            main_frame,
            label_text="処理対象リスト (ファイル名を編集可能)",
            fg_color=Colors.BACKGROUND_PANEL,
            label_fg_color=Colors.adjust_brightness(Colors.BACKGROUND_PANEL),
        )
        self.staged_list_frame.grid(row=0, column=1, pady=0, padx=(5, 0), sticky="nsew")

        # 4. 処理対象ファイル数のラベル
        self.staged_files_label = ctk.CTkLabel(self, text="処理対象ファイル: 0 件")
        self.staged_files_label.grid(row=2, column=0, pady=5, padx=10)

        # 5. 実行ボタン
        self.staged_run_button = ctk.CTkButton(
            self,
            text="出力先を選んで処理実行",
            command=self.run_staged_process,
            state="disabled",
            fg_color=Colors.UI_BASIC,
            hover_color=Colors.adjust_brightness(Colors.UI_BASIC),
            text_color="black",
        )
        self.staged_run_button.grid(row=3, column=0, pady=10, padx=10, sticky="ew")

        # 実行ボタンの状態を初期更新
        self.update_staged_files_label()

    def iconbitmap(self, *args, **kwargs):
        """
        CustomTkinterがアイコンをリセットするのを防ぐためのオーバーライドメソッド。
        """
        # 設定済みのアイコンパスがあれば、引数を無視してそれを適用する
        if hasattr(self, "_custom_icon_path") and self._custom_icon_path:
            try:
                super().iconbitmap(self._custom_icon_path)
            except Exception:
                pass
        else:
            try:
                super().iconbitmap(*args, **kwargs)
            except Exception:
                pass

    def on_close(self) -> None:
        """
        ウィンドウが閉じられるとき（[x]ボタン押下時）の処理。
        Ctrl+Vのバインドを解除し、ウィンドウを破棄します。
        """
        try:
            # メインウィンドウの bind_all を復元
            self.parent_app.bind_all("<Control-v>", lambda e: None)

            # プレビューウィンドウも明示的に破棄する
            if self.preview_popup:
                try:
                    self.preview_popup.destroy()
                except Exception:
                    pass

            self.grab_release()
            self.destroy()
        except Exception as e:
            logger.error(f"DND window close error: {e}")

    def on_merge_toggle(self):
        """「統合」チェックボックス切り替え時のUI制御"""
        # UI上の全Entryウィジェットへの参照を取得
        ui_row_frames = self.staged_list_frame.winfo_children()
        ui_entries = []
        for row_frame in ui_row_frames:
            try:
                bottom_frame = row_frame.winfo_children()[1]
                entry = bottom_frame.winfo_children()[1]
                if isinstance(entry, ctk.CTkEntry):
                    ui_entries.append(entry)
            except IndexError:
                pass

        if self.merge_files_checkbox.get():
            self.staged_list_frame.configure(
                label_text="処理対象リスト (1番目のファイル名が採用されます)"
            )
            new_integrated_name = ""

            # 内部データ (staged_items) をループ
            for i, item in enumerate(self.staged_items):
                if i == 0:
                    # 1番目のアイテム
                    current_name = item["base_name_var"].get()
                    if (
                        not re.match(r"^\d{14}_", current_name)
                        or "貼り付け" in current_name
                    ):
                        now = datetime.datetime.now()
                        time_text = now.strftime("%Y%m%d_%H%M%S")
                        new_integrated_name = f"{time_text}_統合クリップ"
                    else:
                        new_integrated_name = f"{current_name}_統合クリップ"

                    item["base_name_var"].set(new_integrated_name)

                    # 1番目のUIを有効化 + バインド
                    if i < len(ui_entries):
                        ui_entries[i].configure(state="normal")
                        ui_entries[i].bind("<KeyRelease>", self.sync_merge_name)
                else:
                    # 2番目以降のアイテム
                    item["base_name_var"].set(new_integrated_name)
                    # 2番目以降のUIを無効化
                    if i < len(ui_entries):
                        ui_entries[i].configure(state="disabled")
                        ui_entries[i].unbind("<KeyRelease>")
        else:
            self.staged_list_frame.configure(
                label_text="処理対象リスト (ファイル名を編集可能)"
            )

            # 全アイテムのファイル名編集欄を「元の名前」に戻し＆有効化
            for i, item in enumerate(self.staged_items):

                if isinstance(item["data"], Path):
                    original_stem = item["data"].stem
                elif "貼り付け" in item["original_name"]:
                    # 貼り付けの場合、新しいタイムスタンプ名を生成
                    now = datetime.datetime.now()
                    time_text = now.strftime("%Y%m%d_%H%M%S")
                    original_stem = f"{time_text}_{i:02d}_貼り付け"
                else:
                    original_stem = f"file_{i:02d}"  # フォールバック

                item["base_name_var"].set(original_stem)

                # UI上の全Entryを有効化 + バインド解除
                if i < len(ui_entries):
                    ui_entries[i].configure(state="normal")
                    ui_entries[i].unbind("<KeyRelease>")

    def sync_merge_name(self, event=None):
        """
        (統合モード時) 1番目のEntryの変更を、
        他のすべてのアイテムの base_name_var に同期する
        """
        if not self.merge_files_checkbox.get() or not self.staged_items:
            return

        # 1番目の名前を取得
        new_name = self.staged_items[0]["base_name_var"].get()

        # 2番目以降の内部データ(StringVar)に反映
        for item in self.staged_items[1:]:
            item["base_name_var"].set(new_name)

    def add_staged_item(self, item_data, base_name: str, original_name: str) -> bool:
        """
        処理対象アイテムを内部リスト (self.staged_items) とUI (リストフレーム) に
        追加します。

        Args:
            item_data (Path | Image.Image):
                処理対象のデータ（ファイルパスまたはPillowイメージ）。
            base_name (str): 出力ファイル名のベース（編集可能）。
            original_name (str): 元のファイル名または識別子（表示用）。

        Returns:
            bool: アイテムが新しく追加された場合はTrue、
                  (重複などで) スキップされた場合はFalse。
        """
        # 重複チェック (Pathオブジェクトの場合のみ)
        if isinstance(item_data, Path):
            existing_paths = {
                item["data"]
                for item in self.staged_items
                if isinstance(item["data"], Path)
            }
            if item_data in existing_paths:
                logger.warning(f"スキップ (既に追加済み): {original_name}")
                return False

        # 内部リストにアイテム辞書を追加
        item = {
            "data": item_data,
            "base_name_var": ctk.StringVar(value=base_name),
            "original_name": original_name,
            "id": str(len(self.staged_items)),
        }
        self.staged_items.append(item)

        self._create_row_widget(item)
        return True

    def _create_row_widget(self, item):
        """アイテムの行ウィジェットを生成して配置する"""
        row_frame = ctk.CTkFrame(
            self.staged_list_frame,
            fg_color=Colors.blend_colors("#000000", Colors.BACKGROUND_PANEL, 0.60),
        )
        row_frame.item = item

        # 上段: 元ファイル名 + 編集ボタン
        top_sub_frame = ctk.CTkFrame(row_frame, fg_color="transparent")
        top_sub_frame.pack(side="top", fill="x", padx=5, pady=(2, 0))

        name_label = ctk.CTkLabel(
            top_sub_frame,
            text=item["original_name"],
            font=("", 10),
            text_color=Colors.blend_colors("#FFFFFF", Colors.BACKGROUND_PANEL, 0.30),
        )
        name_label.pack(side="left", fill="x", expand=True)

        # 編集ボタン
        item_data = item["data"]
        is_editable = False
        if isinstance(item_data, Image.Image):
            is_editable = True
        elif isinstance(item_data, Path) and item_data.suffix.lower() in [
            ".png",
            ".jpg",
            ".jpeg",
            ".bmp",
        ]:
            is_editable = True

        if is_editable:
            bg = Colors.blend_colors("#000000", Colors.BACKGROUND_PANEL, 0.70)
            edit_btn = ctk.CTkButton(
                top_sub_frame,
                text="変形/Crop",
                width=60,
                height=20,
                font=("", 10),
                fg_color=bg,
                hover_color=Colors.adjust_brightness(bg),
                text_color=Colors.blend_colors(
                    "#FFFFFF", Colors.BACKGROUND_PANEL, 0.30
                ),
                command=lambda i=item: self.open_crop_editor(i),
            )
            edit_btn.pack(side="right", padx=5)

        # 下段: [プレビュー] [エントリー] [削除ボタン]
        bottom_sub_frame = ctk.CTkFrame(row_frame, fg_color="transparent")
        bottom_sub_frame.pack(side="top", fill="x", padx=5, pady=(0, 5))

        # プレビュー兼移動用ハンドル
        preview_label = ctk.CTkLabel(
            bottom_sub_frame,
            text="👁",
            width=30,
            height=28,
            font=("", 16),
            fg_color=Colors.blend_colors("#000000", Colors.BACKGROUND_PANEL, 0.70),
            text_color=Colors.blend_colors("#FFFFFF", Colors.BACKGROUND_PANEL, 0.30),
            corner_radius=6,
            cursor="fleur",
        )
        preview_label.pack(side="left", padx=(0, 5))

        # イベントバインド
        preview_label.bind("<Enter>", lambda e, i=item: self._on_hover_enter(e, i))
        preview_label.bind("<Leave>", self._on_hover_leave)
        preview_label.bind("<Button-1>", lambda e, i=item: self._on_reorder_start(e, i))
        preview_label.bind("<B1-Motion>", self._on_reorder_drag)
        preview_label.bind("<ButtonRelease-1>", self._on_reorder_stop)

        name_entry = ctk.CTkEntry(
            bottom_sub_frame,
            textvariable=item["base_name_var"],
            fg_color=(Colors.BACKGROUND_HOLLOW, Colors.BACKGROUND_DARK_HOLLOW),
        )

        if self.merge_files_checkbox.get():
            if self.staged_items:
                is_first = self.staged_items[0] == item
            else:
                is_first = False
            if is_first:
                name_entry.configure(state="normal")
                name_entry.bind("<KeyRelease>", self.sync_merge_name)
            else:
                name_entry.configure(state="disabled")

        name_entry.pack(side="left", fill="x", expand=True, padx=(0, 5))

        # 削除ボタン (右側)
        delete_btn = ctk.CTkButton(
            bottom_sub_frame,
            text="x",
            width=28,
            height=28,
            fg_color=Colors.LABEL_DENGER,
            hover_color=Colors.adjust_brightness(Colors.LABEL_DENGER, 0.6),
            command=lambda i=item, rf=row_frame: self.remove_staged_item(i, rf),
        )
        delete_btn.pack(side="right")

        row_frame.pack(fill="x", pady=5, padx=5)

    def _refresh_staged_list_ui(self):
        """現在の staged_items の順序に従ってリストを再描画する"""
        for widget in self.staged_list_frame.winfo_children():
            widget.destroy()

        for item in self.staged_items:
            self._create_row_widget(item)

        if self.merge_files_checkbox.get():
            self.on_merge_toggle()

    def remove_staged_item(self, item_to_remove: dict, row_frame: ctk.CTkFrame) -> None:
        """
        指定されたアイテムを内部リストとUIから削除します。
        """
        try:
            # 削除する前に、それが1番目のアイテムだったか確認
            is_first_item = (
                len(self.staged_items) > 0 and self.staged_items[0] == item_to_remove
            )

            self.staged_items.remove(item_to_remove)
            row_frame.destroy()
            self.update_staged_files_label()

            # もし統合モードで1番目のアイテムを削除したら、
            # UIの状態（2番目だったアイテムを編集可能にするなど）を更新する
            if self.merge_files_checkbox.get() and is_first_item:
                self.on_merge_toggle()

        except ValueError:
            logger.error("エラー: 削除対象のアイテムがリストに見つかりません。")

    def handle_drop(self, event: tkinterdnd2.TkinterDnD.DnDEvent) -> None:
        """
        D&Dイベント（ファイル/フォルダドロップ）を処理します。

        Args:
            event (tkinterdnd2.TkinterDnD.DnDEvent): D&Dイベントオブジェクト。
        """
        try:
            file_paths = self.tk.splitlist(event.data)
            added_files_count = 0

            for f in file_paths:
                p = Path(f)
                if p.is_file() and p.suffix.lower() in SUPPORTED_EXTENSIONS:
                    if self.add_staged_item(p, p.stem, p.name):
                        added_files_count += 1
                elif p.is_dir():
                    for ext in SUPPORTED_EXTENSIONS:
                        for child_file in p.glob(f"**/*{ext}"):
                            if self.add_staged_item(
                                child_file, child_file.stem, child_file.name
                            ):
                                added_files_count += 1
            self.update_staged_files_label()
            if added_files_count > 0:
                self.parent_app.status_label.configure(
                    text=f"{added_files_count}件のファイルを追加 (D&D)"
                )

        except Exception as e:
            messagebox.showerror(
                "ドロップエラー", f"ファイルの処理に失敗しました: {e}", parent=self
            )

    def handle_paste(self, event=None) -> None:
        """
        ペーストイベント (Ctrl+V) を処理します。
        クリップボード上の画像、またはファイルパス（テキスト）に対応します。

        Args:
            event (None): イベントオブジェクト (使用しない)。
        """
        try:
            # 1. 画像の貼り付けを試みる
            pil_image = ImageGrab.grabclipboard()

            if pil_image:
                now = datetime.datetime.now()
                base_name = f"{now.strftime('%Y%m%d_%H%M%S')}_貼り付け"
                original_name = (
                    "クリップボードの画像 " + f"({pil_image.width}x{pil_image.height})"
                )

                self.add_staged_item(pil_image, base_name, original_name)
                self.update_staged_files_label()
                self.parent_app.status_label.configure(
                    text=f"クリップボードの画像を追加 ( {base_name} )"
                )

            else:
                # 2. 画像でない場合、テキスト (ファイルパス) として取得を試みる
                clipboard_content = self.clipboard_get()
                file_paths = [
                    line.strip().strip('"')
                    for line in re.split(r"[\n\t]", clipboard_content)
                    if line.strip()
                ]

                added_files_count = 0
                for f in file_paths:
                    p = Path(f)
                    if p.is_file() and p.suffix.lower() in SUPPORTED_EXTENSIONS:
                        if self.add_staged_item(p, p.stem, p.name):
                            added_files_count += 1
                if added_files_count > 0:
                    self.update_staged_files_label()
                    self.parent_app.status_label.configure(
                        text=f"{added_files_count}件のファイルパスを追加 (貼付)"
                    )

        except Exception as e:
            if "CLIPBOARD" in str(e):
                self.parent_app.status_label.configure(
                    text="クリップボードが空か、対応形式ではありません。",
                    text_color=Colors.LABEL_WARNING,
                )
            else:
                messagebox.showerror(
                    "貼付エラー",
                    f"クリップボードの処理に失敗しました: {e}",
                    parent=self,
                )

    def update_staged_files_label(self) -> None:
        """
        処理対象リストの件数をラベルに反映し、
        実行ボタンの有効/無効状態を切り替えます。
        """
        count = len(self.staged_items)
        self.staged_files_label.configure(text=f"処理対象ファイル: {count} 件")
        if (
            count > 0
            and self.parent_app.font_path
            and Path(self.parent_app.font_path).is_file()
        ):
            self.staged_run_button.configure(state="normal")
        else:
            self.staged_run_button.configure(state="disabled")

    # --- プレビュー制御メソッド (ウィンドウ再利用版) ---
    def _ensure_preview_popup(self):
        """プレビューウィンドウが存在しなければ作成し、あれば返す"""
        if self.preview_popup is None or not self.preview_popup.winfo_exists():
            self.preview_popup = ctk.CTkToplevel(self)
            self.preview_popup.wm_overrideredirect(True)
            self.preview_popup.wm_attributes("-topmost", True)

            # 初期状態は非表示
            self.preview_popup.withdraw()

            # 画像表示用ラベル
            self.preview_label = ctk.CTkLabel(self.preview_popup, text="")
            self.preview_label.pack()

    def _perform_preview_show(self, item):
        """予約実行されるプレビュー表示の実処理"""
        # 実行の瞬間に、まだそのアイテム上にいるか最終確認
        if self.hovering_item != item:
            return

        thumb = self._get_thumbnail(item)

        # サムネイル生成（重い処理）の間にマウスが外れていないか再度確認
        if self.hovering_item != item:
            return

        if not thumb:
            return

        # ウィンドウの準備
        self._ensure_preview_popup()

        # 画像を更新
        self.preview_label.configure(image=thumb)

        # 現在のマウス位置を取得 (event.x_root ではなく現在のポインタ位置を使う)
        x, y = self.winfo_pointerxy()
        x += 20
        y += 20

        # 画面はみ出し対策
        sw = self.winfo_screenwidth()
        sh = self.winfo_screenheight()
        if x + 300 > sw:
            x -= 320
        if y + 300 > sh:
            y -= 320

        self.preview_popup.geometry(f"+{x}+{y}")

        # 表示して最前面へ
        self.preview_popup.deiconify()
        self.preview_popup.lift()

    def _close_preview(self):
        """プレビューウィンドウを非表示にする (破棄はしない)"""
        if self.preview_popup and self.preview_popup.winfo_exists():
            self.preview_popup.withdraw()

    def _on_hover_enter(self, event, item):
        """マウスオーバー時にプレビュー表示をスケジュールする"""
        self.hovering_item = item

        # 既存のタイマーがあればキャンセル (連打防止)
        if self.preview_timer:
            self.after_cancel(self.preview_timer)
            self.preview_timer = None

        # 0.2秒 (200ms) 後にプレビュー処理を実行するように予約
        # これにより、素早く通り過ぎただけの時は処理が走らない
        self.preview_timer = self.after(200, lambda: self._perform_preview_show(item))

    def _on_hover_leave(self, event):
        """マウスが離れたらプレビューをキャンセル・非表示にする"""
        self.hovering_item = None

        # 予約されていた表示処理をキャンセル
        if self.preview_timer:
            self.after_cancel(self.preview_timer)
            self.preview_timer = None

        self._close_preview()

    def _on_reorder_start(self, event, item):
        """ドラッグ開始時はプレビューを隠す"""
        self.drag_source_item = item
        self._close_preview()

    def _on_reorder_drag(self, event):
        """ドラッグ中: 並び替え"""
        if not self.drag_source_item:
            return
        x, y = event.x_root, event.y_root
        target_widget = self.winfo_containing(x, y)
        if not target_widget:
            return

        target_row = None
        curr = target_widget
        for _ in range(10):
            if hasattr(curr, "item"):
                target_row = curr
                break
            try:
                parent_name = curr.winfo_parent()
                if not parent_name:
                    break
                curr = self._nametowidget(parent_name)
            except (KeyError, TypeError, AttributeError):
                break

        if target_row and hasattr(target_row, "item"):
            target_item = target_row.item
            if (
                target_item != self.drag_source_item
                and target_item in self.staged_items
            ):
                src_idx = self.staged_items.index(self.drag_source_item)
                dst_idx = self.staged_items.index(target_item)
                self.staged_items.pop(src_idx)
                self.staged_items.insert(dst_idx, self.drag_source_item)
                self._refresh_staged_list_ui()

    def _on_reorder_stop(self, event):
        """ドロップ (終了)"""
        self.drag_source_item = None

    def _get_thumbnail(self, item):
        item_id = id(item)
        if item_id in self.preview_image_cache:
            return self.preview_image_cache[item_id]

        data = item["data"]
        pil_img = None
        max_size = (300, 300)

        try:
            if isinstance(data, Image.Image):
                pil_img = data.copy()
            elif isinstance(data, Path):
                if data.suffix.lower() == ".pdf":
                    doc = fitz.open(data)
                    pix = doc[0].get_pixmap(dpi=72)
                    pil_img = Image.open(io.BytesIO(pix.tobytes("png")))
                    doc.close()
                elif data.suffix.lower() in [".png", ".jpg", ".jpeg", ".bmp"]:
                    pil_img = Image.open(data)

            if pil_img:
                pil_img.thumbnail(max_size)
                ctk_img = ctk.CTkImage(
                    light_image=pil_img, dark_image=pil_img, size=pil_img.size
                )
                self.preview_image_cache[item_id] = ctk_img
                return ctk_img
        except Exception as e:
            logger.error(f"Thumbnail error: {e}")
        return None

    def open_crop_editor(self, item):
        def on_save(new_pil_image):
            item["data"] = new_pil_image
            item_id = id(item)
            if item_id in self.preview_image_cache:
                del self.preview_image_cache[item_id]
            messagebox.showinfo("完了", "画像の変形を適用しました。", parent=self)

        PerspectiveCropEditor(self, item["data"], on_save)

    def run_staged_process(self) -> None:
        if not self.staged_items:
            messagebox.showinfo(
                "情報", "処理対象のファイルが指定されていません。", parent=self
            )
            return

        # 1. 親アプリから設定情報を取得
        font_path = self.parent_app.font_path
        if not font_path or not Path(font_path).is_file():
            self.parent_app.status_label.configure(
                text="エラー: config.iniで有効なフォントパスが指定されていません。",
                text_color=Colors.LABEL_WARNING,
            )
            return

        config_data = self.parent_app.config_data
        key_rect_tuple = config_data.get("key_rect", (0, 0, 0, 0))
        paper_width = self.parent_app.paper_width
        paper_height = self.parent_app.paper_height
        enable_tesseract = config_data.get("enable_tesseract_ocr", False)

        # Ollama設定を取得
        ollama_model = config_data.get("ollama_model", "")
        ollama_api_url = config_data.get("ollama_api_url", "")

        flatten_ink = config_data.get("flatten_ink_annotations", True)
        paper_size_str = self.parent_app.config_data.get("paper_size", "A4")

        # 2. 出力先フォルダを選択
        dest_folder = filedialog.askdirectory(
            title="出力先フォルダを選択してください", parent=self
        )
        if not dest_folder:
            return
        dest_path = Path(dest_folder)

        same_folder_detected = False
        for item in self.staged_items:
            # 入力元と出力先が同じパスかどうかを確認
            if isinstance(item["data"], Path) and item["data"].parent == dest_path:
                same_folder_detected = True
                break

        if same_folder_detected:
            # エラーで弾くのではなく、確認ダイアログを表示して続行可否をユーザーに委ねる
            if not messagebox.askyesno(
                "確認",
                "出力先に入力元と同じフォルダが含まれています。\n"
                "同名のファイルが存在する場合、元のファイルは正規化後のPDFで上書きされます。\n\n"
                "処理を続行しますか？",
                icon="warning",
                parent=self,
            ):
                return

        # 3. メタデータをUIから取得
        index_key_raw = self.index_key_combo.get()
        index_key_to_embed = ""
        text_color = None
        if index_key_raw != "（未選択）":
            index_key_to_embed = index_key_raw
            key_colors_dict = config_data.get("key_colors", {})
            hex_color = key_colors_dict.get(index_key_raw.lower())
            if hex_color:
                text_color = hex_to_rgb_tuple(hex_color)
        comment_to_embed = self.comment_textbox.get("1.0", "end-1c").strip()

        # 引用Keyリストを取得 ([[Key: Title]]形式に対応)
        cited_keys_str = self.cited_keys_entry.get("1.0", "end-1c").strip()
        cited_keys_list = []

        refs_qr_size_pt = self.parent_app.config_data.get("refs_qr_size", 75)

        # Key (14桁以上の数字) を抽出するための正規表現
        key_regex = re.compile(r"(\d{14,})")
        if cited_keys_str:
            parts = re.split(r"[,\n]+", cited_keys_str)
            for part in parts:
                part_cleaned = part.strip()
                if not part_cleaned:
                    continue

                # 2. 各部分から Key (14桁以上の数字) を検索
                match = key_regex.search(part_cleaned)

                if match:
                    # 3. 見つかったKeyのみをリストに追加
                    extracted_key = match.group(1)
                    if extracted_key not in cited_keys_list:  # 重複防止
                        cited_keys_list.append(extracted_key)

        raw_sist_author = self.sist_author_entry.get().strip()
        raw_sist_title = self.sist_title_entry.get().strip()
        raw_sist_site = self.sist_site_entry.get().strip()
        raw_sist_date = self.sist_date_entry.get().strip()

        sist_view_date = datetime.datetime.now().strftime("%Y-%m-%d")
        sist_string_formal = None
        sist_string_readable = None

        if raw_sist_author or raw_sist_title or raw_sist_site or raw_sist_date:
            sist_author = raw_sist_author or "（著者不明）"
            sist_title = raw_sist_title or "（タイトル不明）"
            sist_site = raw_sist_site or "（サイト名不明）"
            sist_date = raw_sist_date or "（更新日不明）"
            sist_string_formal = (
                f"{sist_author} . “{sist_title}” . {sist_site} . "
                f"{sist_date} . (参照 {sist_view_date})"
            )
            sist_string_readable = (
                f"著者:{sist_author}\n\nページ名:\n“{sist_title}”\n\n"
                f"サイト名/出典:\n{sist_site}\n\n"
                f"更新日:{sist_date} (参照:{sist_view_date})"
            )

        # 4. 自動タグ付け用キーワードの生成 ---
        extra_keywords = []

        # 4-1. 書誌情報が入力されている場合 -> Synapsen:Source
        if raw_sist_author or raw_sist_title or raw_sist_site or raw_sist_date:
            extra_keywords.append("Synapsen:Source")

        # 4-2. WebClipチェックボックスがONの場合 -> Synapsen:WebClip
        if self.is_webclip_var.get():
            extra_keywords.append("Synapsen:WebClip")

        # 5. 処理対象リストを作成
        items_to_process = []

        # ファイル名に使用できない文字を置換する正規表現
        invalid_chars_pattern = re.compile(r'[\\/:\*\?"<>\|]')

        for item in self.staged_items:
            # 元のファイル名を取得
            base_name_raw = item["base_name_var"].get().strip()

            # サニタイズ処理を実行
            base_name = invalid_chars_pattern.sub("_", base_name_raw)

            if not base_name:
                messagebox.showerror(
                    "ファイル名エラー",
                    f"ファイル名が空です (元の名前: {item['original_name']})",
                    parent=self,
                )
                return

            # もし置換が発生したら、UIの表示 (StringVar) にも反映する
            if base_name != base_name_raw:
                item["base_name_var"].set(base_name)

            items_to_process.append((item["data"], base_name, type(item["data"])))

        # 6. 「統合」チェックボックスの状態で処理を分岐
        is_merge_mode = self.merge_files_checkbox.get()
        temp_dir = None

        try:
            temp_dir = Path(tempfile.mkdtemp(prefix="synapsen_dnd_", dir=dest_path))

            if not is_merge_mode:
                # --- [分岐 A: 個別処理] ---
                self.run_pipeline_individual(
                    items_to_process,
                    dest_path,
                    temp_dir,
                    font_path,
                    paper_width,
                    paper_height,
                    enable_tesseract,
                    ollama_model,
                    ollama_api_url,
                    paper_size_str,
                    key_rect_tuple,
                    index_key_to_embed,
                    text_color,
                    comment_to_embed,
                    sist_string_formal,
                    sist_string_readable,
                    cited_keys_list,
                    refs_qr_size_pt,
                    flatten_ink,
                    extra_keywords=extra_keywords,
                )
            else:
                # --- [分岐 B: 統合処理] ---
                self.run_pipeline_merge(
                    items_to_process,
                    dest_path,
                    temp_dir,
                    font_path,
                    paper_width,
                    paper_height,
                    enable_tesseract,
                    ollama_model,
                    ollama_api_url,
                    paper_size_str,
                    key_rect_tuple,
                    index_key_to_embed,
                    text_color,
                    comment_to_embed,
                    sist_string_formal,
                    sist_string_readable,
                    cited_keys_list,
                    refs_qr_size_pt,
                    flatten_ink,
                    extra_keywords=extra_keywords,
                )
            self.on_close()

        except Exception as e:
            messagebox.showerror(
                "処理エラー", f"処理中にエラーが発生しました:\n{e}", parent=self
            )
            self.parent_app.status_label.configure(text=f"エラーが発生しました: {e}")
        finally:
            if temp_dir and temp_dir.exists():
                try:
                    shutil.rmtree(temp_dir)
                except Exception as e:
                    logger.error(f"警告: 一時フォルダの削除に失敗しました: {e}")

    def run_pipeline_individual(
        self,
        items_to_process,
        dest_path,
        temp_dir,
        font_path,
        paper_width,
        paper_height,
        enable_tesseract,
        ollama_model,
        ollama_api_url,
        paper_size_str,
        key_rect_tuple,
        index_key_to_embed,
        text_color,
        comment_to_embed,
        sist_string_formal,
        sist_string_readable,
        cited_keys_list,
        refs_qr_size_pt,
        flatten_ink=True,
        extra_keywords=None,
    ):
        total_files = len(items_to_process)
        for i, item_tuple in enumerate(items_to_process):
            item_data, base_name, original_type = item_tuple
            status_prefix = f"処理中 ({i+1}/{total_files}):"
            self.parent_app.status_label.configure(text=f"{status_prefix} {base_name}")
            self.parent_app.update_idletasks()

            final_output_pdf = dest_path / f"{base_name}.pdf"
            temp_converted_pdf = temp_dir / f"conv_{base_name}.pdf"
            temp_flattened_pdf = temp_dir / f"flat_{base_name}.pdf"
            path_to_flatten: Path

            if isinstance(item_data, Path) and item_data.suffix.lower() == ".md":
                try:
                    temp_md_pdf = temp_dir / f"md_{base_name}.pdf"
                    convert_document_to_pdf(item_data, temp_md_pdf, paper_size_str)
                    item_data = temp_md_pdf
                except Exception as e:
                    logger.warning(f"警告: {base_name} のMarkdown変換に失敗: {e}")
                    messagebox.showerror(
                        "Markdown変換エラー",
                        f"{base_name} の変換に失敗しました:\n{e}",
                        parent=self,
                    )
                    continue

            if isinstance(item_data, Path):
                if item_data.suffix.lower() != ".pdf":
                    convert_image_to_pdf(item_data, temp_converted_pdf)
                    path_to_flatten = temp_converted_pdf
                else:
                    path_to_flatten = item_data
            else:
                convert_pil_image_to_pdf(item_data, temp_converted_pdf)
                path_to_flatten = temp_converted_pdf

            high_fidelity_flatten(
                str(path_to_flatten),
                str(temp_flattened_pdf),
                font_path,
                flatten_ink=flatten_ink,
            )

            # --- 3: 正規化 ---
            self.parent_app.status_label.configure(
                text=f"{status_prefix} 正規化中: {base_name}"
            )
            self.parent_app.update_idletasks()
            normalize_pdf_to_papersize(
                str(temp_flattened_pdf),
                str(final_output_pdf),
                paper_width,
                paper_height,
                target_format=paper_size_str,
            )

            # --- 4: OCR ---
            # OCR設定の判定 (ファイル名末尾によるLocal LLM判定)
            use_ollama = False
            if isinstance(base_name, str):
                lower_name = base_name.lower()
                if lower_name.endswith("_hand") or lower_name.endswith("_llm"):
                    use_ollama = True

            current_ocr_engine = "tesseract"
            should_run_ocr = enable_tesseract

            if use_ollama:
                if ollama_model:
                    current_ocr_engine = "ollama"
                    should_run_ocr = True
                else:
                    logger.warning(
                        f"Ollamaフラグ({base_name})を検出しましたが、"
                        "configにモデル設定がないためOCRをスキップします。"
                    )
                    should_run_ocr = False

            self.parent_app.status_label.configure(
                text=f"{status_prefix} OCR埋込処理中({current_ocr_engine}): {base_name}"
            )
            self.parent_app.update_idletasks()

            embed_ocr_text_in_pdf(
                str(final_output_pdf),
                should_run_ocr,
                font_path,
                ocr_engine=current_ocr_engine,
                ollama_config={
                    "model": ollama_model,
                    "url": ollama_api_url,
                },
            )

            # --- 5: メタデータ追記 ---
            self.parent_app.status_label.configure(
                text=f"{status_prefix} メタデータを追記中: {base_name}"
            )
            self.parent_app.update_idletasks()

            add_metadata_to_clip(
                str(final_output_pdf),
                font_path,
                paper_width,
                paper_height,
                key_rect_tuple,
                index_key_to_embed,
                text_color,
                comment_to_embed,
                sist_string_formal,
                sist_string_readable,
                base_name=base_name,
                cited_keys_list=cited_keys_list,
                refs_qr_size_pt=refs_qr_size_pt,
                extra_keywords=extra_keywords,
            )

        messagebox.showinfo(
            "完了",
            f"{total_files}個のファイルにメタデータを埋め込み、処理が完了しました。",
            parent=self,
        )
        self.parent_app.status_label.configure(text="処理が完了しました。")

    def run_pipeline_merge(
        self,
        items_to_process,
        dest_path,
        temp_dir,
        font_path,
        paper_width,
        paper_height,
        enable_tesseract,
        ollama_model,  # Arg Added
        ollama_api_url,  # Arg Added
        paper_size_str,
        key_rect_tuple,
        index_key_to_embed,
        text_color,
        comment_to_embed,
        sist_string_formal,
        sist_string_readable,
        cited_keys_list,
        refs_qr_size_pt,
        flatten_ink=True,
        extra_keywords=None,
    ):
        """
        ファイル統合処理パイプライン
        """

        # 統合後のファイル名は、リストの最初のアイテムから取得
        if not items_to_process:
            return
        merged_base_name = items_to_process[0][1].strip()
        if not merged_base_name:
            messagebox.showerror(
                "ファイル名エラー", "統合後のファイル名が空です。", parent=self
            )
            return

        final_output_pdf = dest_path / f"{merged_base_name}.pdf"

        # 正規化済みPDFを格納する一時サブフォルダ
        normalized_parts_dir = temp_dir / "normalized_parts"
        normalized_parts_dir.mkdir()

        total_files = len(items_to_process)
        normalized_pdf_paths = []  # 連結するPDFのパスリスト

        for i, item_tuple in enumerate(items_to_process):
            item_data, base_name, original_type = item_tuple

            status_prefix = f"統合準備中 ({i+1}/{total_files}):"

            # self.staged_items はD&Dウィンドウで管理されている元のリスト
            # ここから元のファイル名/表示名を取得してOllama判定に使用する
            # (統合時は出力ファイル名 base_name が統一されてしまうため)
            item_original_name = self.staged_items[i]["original_name"]

            self.parent_app.status_label.configure(
                text=f"{status_prefix} {item_original_name}"
            )
            self.parent_app.update_idletasks()

            # 各ステップの一時ファイル
            temp_converted_pdf = temp_dir / f"conv_{i}_{base_name}.pdf"
            temp_flattened_pdf = temp_dir / f"flat_{i}_{base_name}.pdf"
            file_name = f"part_{i:03d}_{base_name}.pdf"
            normalized_part_pdf = normalized_parts_dir / file_name

            path_to_flatten: Path

            # --- 1-A: MD -> PDF ---
            if isinstance(item_data, Path) and item_data.suffix.lower() == ".md":
                self.parent_app.status_label.configure(
                    text=f"{status_prefix} MD->PDF変換: {item_original_name}"
                )
                self.parent_app.update_idletasks()
                try:
                    temp_md_pdf = temp_dir / f"md_{i}_{base_name}.pdf"
                    convert_document_to_pdf(item_data, temp_md_pdf, paper_size_str)
                    item_data = temp_md_pdf
                except Exception as e:
                    logger.error(
                        f"警告: {item_original_name} のMarkdown変換に失敗: {e}"
                    )
                    continue

            # --- 1-B: 画像/PIL -> PDF ---
            if isinstance(item_data, Path):
                if item_data.suffix.lower() != ".pdf":
                    convert_image_to_pdf(item_data, temp_converted_pdf)
                    path_to_flatten = temp_converted_pdf
                else:
                    path_to_flatten = item_data
            else:
                convert_pil_image_to_pdf(item_data, temp_converted_pdf)
                path_to_flatten = temp_converted_pdf

            # --- 2: フラット化 ---
            self.parent_app.status_label.configure(
                text=f"{status_prefix} フラット化中: {item_original_name}"
            )
            self.parent_app.update_idletasks()
            high_fidelity_flatten(
                str(path_to_flatten),
                str(temp_flattened_pdf),
                font_path,
                flatten_ink=flatten_ink,
            )

            # --- 3: 正規化 (出力先を normalized_part_pdf に) ---
            self.parent_app.status_label.configure(
                text=f"{status_prefix} 正規化中: {item_original_name}"
            )
            self.parent_app.update_idletasks()
            normalize_pdf_to_papersize(
                str(temp_flattened_pdf),
                str(normalized_part_pdf),
                paper_width,
                paper_height,
                target_format=paper_size_str,
            )

            # --- 4: OCR (normalized_part_pdf に対して実行) ---
            # OCR設定の判定 (元のファイル名末尾によるLocal LLM判定)
            use_ollama = False
            # 統合モードでは base_name は全て同じ名前になっているため、
            # item_original_name (D&Dされた元のファイル名) を使用して判定する
            if isinstance(item_original_name, str):
                p_original = Path(item_original_name)
                stem_original = p_original.stem.lower()
                if stem_original.endswith("_hand") or stem_original.endswith("_llm"):
                    use_ollama = True

            current_ocr_engine = "tesseract"
            should_run_ocr = enable_tesseract

            if use_ollama:
                if ollama_model:
                    current_ocr_engine = "ollama"
                    should_run_ocr = True
                else:
                    logger.warning(
                        f"Ollamaフラグ({item_original_name})を検出しましたが、"
                        "configにモデル設定がないためOCRをスキップします。"
                    )
                    should_run_ocr = False

            self.parent_app.status_label.configure(
                text=(
                    f"{status_prefix} OCR埋込処理中({current_ocr_engine}): "
                    f"{item_original_name}"
                )
            )
            self.parent_app.update_idletasks()

            embed_ocr_text_in_pdf(
                str(normalized_part_pdf),
                should_run_ocr,
                font_path,
                ocr_engine=current_ocr_engine,
                ollama_config={
                    "model": ollama_model,
                    "url": ollama_api_url,
                },
            )

            # 連結リストに追加
            normalized_pdf_paths.append(normalized_part_pdf)

        # --- ループ終了後 ---

        if not normalized_pdf_paths:
            messagebox.showerror(
                "エラー", "統合できるファイルがありませんでした。", parent=self
            )
            return

        # --- 5: PDF連結 ---
        self.parent_app.status_label.configure(
            text=f"全 {len(normalized_pdf_paths)} ファイルを連結中..."
        )
        self.parent_app.update_idletasks()

        writer = PdfWriter()
        for pdf_path in normalized_pdf_paths:
            try:
                reader = PdfReader(str(pdf_path))
                for page in reader.pages:
                    writer.add_page(page)
            except Exception as e:
                logger.error(f"警告: {pdf_path.name} の連結に失敗: {e}")

        # 連結したPDFを一時ファイル (final_output_pdf の場所) に保存
        with open(final_output_pdf, "wb") as f:
            writer.write(f)
        writer.close()

        # --- 6: メタデータ追記 (連結後のPDFに対して1回だけ実行) ---
        self.parent_app.status_label.configure(text="メタデータを追記中...")
        self.parent_app.update_idletasks()

        add_metadata_to_clip(
            str(final_output_pdf),
            font_path,
            paper_width,
            paper_height,
            key_rect_tuple,
            index_key_to_embed,
            text_color,
            comment_to_embed,
            sist_string_formal,
            sist_string_readable,
            base_name=base_name,
            cited_keys_list=cited_keys_list,
            refs_qr_size_pt=refs_qr_size_pt,
            extra_keywords=extra_keywords,
        )

        messagebox.showinfo(
            "完了",
            f"{total_files}個のファイルを1つのPDFに統合し、処理が完了しました。",
            parent=self,
        )
        self.parent_app.status_label.configure(
            text="処理が完了しました。", text_color="gray"
        )

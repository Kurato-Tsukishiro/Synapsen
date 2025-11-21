import customtkinter as ctk
from tkinter import filedialog, messagebox
from pathlib import Path
import datetime
import re
import tkinterdnd2
from PIL import ImageGrab
import tempfile
import shutil
from pypdf import PdfReader, PdfWriter

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
    embed_ocr_text_in_pdf
)

import logging
logger = logging.getLogger(__name__)

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
    ".odt"
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
                "DND初期化エラー", f"tkinterdnd2 の初期化に失敗しました: {e}", parent=self)
            self.destroy()
            return

        self.parent_app = parent_app  # メインアプリ(Synapsen_Normalisierer)
        self.staged_items = []

        self.title("D&D/ペーストで正規化")
        self.geometry("1000x800")

        if self.parent_app.icon_path:
            try:
                self.iconbitmap(default=str(self.parent_app.icon_path))
            except Exception as e:
                logger.error(f"Icon set error (DND Window): {e}")

        self.protocol("WM_DELETE_WINDOW", self.on_close)
        self.transient(parent_app)
        self.grab_set()

        # --- [UI定義] ---
        self.grid_rowconfigure(1, weight=1)  # 2カラムのメインフレーム領域
        self.grid_columnconfigure(0, weight=1)

        # 1. ドラッグ＆ドロップ ゾーン
        self.drop_target_frame = ctk.CTkFrame(
            self, height=100, fg_color="gray25")
        self.drop_target_frame.grid(
            row=0, column=0, pady=10, padx=10, sticky="ew")

        self.drop_target_label = ctk.CTkLabel(
            self.drop_target_frame,
            text=(
                "ここにファイル (PDF, JPG, PNG, MD) を\nドラッグ＆ドロップしてください\n\n" +
                "(または Ctrl+V でスクリーンショットを貼付)"
            ),
            text_color="gray70"
        )
        self.drop_target_label.place(relx=0.5, rely=0.5, anchor="center")

        # D&Dイベントのバインド
        self.drop_target_frame.drop_target_register(tkinterdnd2.DND_FILES)
        self.drop_target_frame.dnd_bind('<<Drop>>', self.handle_drop)
        # ペーストイベントのバインド
        self.bind_all("<Control-v>", self.handle_paste)

        main_frame = ctk.CTkFrame(self, fg_color="transparent")
        main_frame.grid(row=1, column=0, pady=0, padx=10, sticky="nsew")
        main_frame.grid_rowconfigure(0, weight=1)
        main_frame.grid_columnconfigure(0, weight=1)  # 左カラム (メタデータ)
        main_frame.grid_columnconfigure(1, weight=1)  # 右カラム (リスト)
        # --- [追加ここまで] ---

        # 2. メタデータ入力フレーム (左カラム)
        meta_frame = ctk.CTkFrame(main_frame)
        meta_frame.grid(row=0, column=0, pady=0, padx=(0, 5), sticky="nsew")

        # 2-1 IndexKey 選択
        ctk.CTkLabel(
            meta_frame, text="IndexKey (PDF 1ページ目に埋込):", anchor="w"
            ).pack(pady=(10, 0), padx=10, fill="x")

        key_options = self.parent_app.config_data.get(
                'commonplace_keys_options', [])

        self.index_key_combo = ctk.CTkComboBox(
            meta_frame,
            values=["（未選択）"] + key_options
        )
        self.index_key_combo.set("（未選択）")
        self.index_key_combo.pack(pady=5, padx=10, fill="x")

        # 2-2. コメント入力
        ctk.CTkLabel(
            meta_frame, text="コメント (PDF 最終ページに埋込):", anchor="w"
            ).pack(pady=(5, 0), padx=10, fill="x")
        self.comment_textbox = ctk.CTkTextbox(meta_frame, height=80)
        self.comment_textbox.pack(
            pady=5, padx=10, fill="both", expand=True)

        # 2-3. 書誌情報UI
        ctk.CTkLabel(
            meta_frame, text="書誌情報 (任意入力)", anchor="w"
            ).pack(pady=(10, 0), padx=10, fill="x")

        sist_frame = ctk.CTkFrame(meta_frame)
        sist_frame.pack(pady=(0, 5), padx=5, fill="x")
        sist_frame.grid_columnconfigure(1, weight=1)

        ctk.CTkLabel(
            sist_frame, text="著者名:"
            ).grid(row=0, column=0, padx=5, pady=5, sticky="w")
        self.sist_author_entry = ctk.CTkEntry(
            sist_frame, placeholder_text="（任意）")
        self.sist_author_entry.grid(
            row=0, column=1, padx=5, pady=5, sticky="ew")

        ctk.CTkLabel(
            sist_frame, text="ページ名:"
            ).grid(row=1, column=0, padx=5, pady=5, sticky="w")
        self.sist_title_entry = ctk.CTkEntry(
            sist_frame, placeholder_text="（任意）")
        self.sist_title_entry.grid(
            row=1, column=1, padx=5, pady=5, sticky="ew")

        ctk.CTkLabel(
            sist_frame, text="サイト名/出典:"
            ).grid(row=2, column=0, padx=5, pady=5, sticky="w")
        self.sist_site_entry = ctk.CTkEntry(
            sist_frame, placeholder_text="（任意）")
        self.sist_site_entry.grid(
            row=2, column=1, padx=5, pady=5, sticky="ew")

        ctk.CTkLabel(
            sist_frame, text="更新日:"
            ).grid(row=3, column=0, padx=5, pady=5, sticky="w")
        self.sist_date_entry = ctk.CTkEntry(
            sist_frame, placeholder_text="（任意, YYYY-MM-DD）")
        self.sist_date_entry.grid(
            row=3, column=1, padx=5, pady=5, sticky="ew")

        # 2-4. 引用Key入力欄
        ctk.CTkLabel(
            meta_frame, text="引用元Key (カンマ区切り または 改行区切り):", anchor="w"
            ).pack(pady=(10, 0), padx=10, fill="x")
        self.cited_keys_entry = ctk.CTkTextbox(
            meta_frame,
            height=40
        )
        self.cited_keys_entry.pack(
            pady=5, padx=10,
            fill="both", expand=True
        )

        # 2-5. 統合チェックボックス
        self.merge_files_checkbox = ctk.CTkCheckBox(
            meta_frame,
            text="処理対象ファイルを1つのノート（PDF）に統合する",
            command=self.on_merge_toggle
        )
        self.merge_files_checkbox.pack(pady=10, padx=10, anchor="w")

        # 3. 処理対象リスト (スクロールフレーム) (右カラム)
        self.staged_list_frame = ctk.CTkScrollableFrame(
            main_frame, label_text="処理対象リスト (ファイル名を編集可能)")
        # .pack() -> .grid()
        self.staged_list_frame.grid(
            row=0, column=1, pady=0, padx=(5, 0), sticky="nsew")

        # 4. 処理対象ファイル数のラベル (ウィンドウ下部)
        self.staged_files_label = ctk.CTkLabel(
            self, text="処理対象ファイル: 0 件")
        # .pack() -> .grid()
        self.staged_files_label.grid(row=2, column=0, pady=5, padx=10)

        # 5. 実行ボタン (ウィンドウ下部)
        self.staged_run_button = ctk.CTkButton(
            self,
            text="出力先を選んで処理実行",
            command=self.run_staged_process,
            state="disabled"
        )
        # .pack() -> .grid()
        self.staged_run_button.grid(
            row=3, column=0, pady=10, padx=10, sticky="ew")

        # 実行ボタンの状態を初期更新
        self.update_staged_files_label()

    def on_close(self) -> None:
        """
        ウィンドウが閉じられるとき（[x]ボタン押下時）の処理。
        Ctrl+Vのバインドを解除し、ウィンドウを破棄します。
        """
        try:
            # メインウィンドウの bind_all を復元
            self.parent_app.bind_all("<Control-v>", lambda e: None)
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
                entry = row_frame.winfo_children()[1]  # 2番目のウィジェット(Entry)
                if isinstance(entry, ctk.CTkEntry):
                    ui_entries.append(entry)
            except IndexError:
                pass

        if self.merge_files_checkbox.get():
            # --- 統合がONの場合 ---
            self.staged_list_frame.configure(
                label_text="処理対象リスト (1番目のファイル名が採用されます)")

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
                        time_text = now.strftime('%Y%m%d_%H%M%S')
                        new_integrated_name = f"{time_text}_統合クリップ"
                    else:
                        new_integrated_name = f"{current_name}_統合クリップ"

                    item["base_name_var"].set(new_integrated_name)

                    # 1番目のUIを有効化 + バインド
                    if i < len(ui_entries):
                        ui_entries[i].configure(state="normal")
                        ui_entries[i].bind(
                            "<KeyRelease>", self.sync_merge_name)
                else:
                    # 2番目以降のアイテム
                    item["base_name_var"].set(new_integrated_name)
                    # 2番目以降のUIを無効化
                    if i < len(ui_entries):
                        ui_entries[i].configure(state="disabled")
                        ui_entries[i].unbind("<KeyRelease>")
        else:
            # --- 統合がOFFの場合 ---
            self.staged_list_frame.configure(
                label_text="処理対象リスト (ファイル名を編集可能)")

            # 全アイテムのファイル名編集欄を「元の名前」に戻し＆有効化
            for i, item in enumerate(self.staged_items):

                if isinstance(item['data'], Path):
                    original_stem = item['data'].stem
                elif "貼り付け" in item['original_name']:
                    # 貼り付けの場合、新しいタイムスタンプ名を生成
                    now = datetime.datetime.now()
                    time_text = now.strftime('%Y%m%d_%H%M%S')
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
        if (not self.merge_files_checkbox.get() or
                not self.staged_items):
            return

        # 1番目の名前を取得
        new_name = self.staged_items[0]["base_name_var"].get()

        # 2番目以降の内部データ(StringVar)に反映
        for item in self.staged_items[1:]:
            item["base_name_var"].set(new_name)

    def add_staged_item(
            self, item_data, base_name: str, original_name: str) -> bool:
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
                item['data']
                for item in self.staged_items
                if isinstance(item['data'], Path)
            }
            if item_data in existing_paths:
                logger.warning(f"スキップ (既に追加済み): {original_name}")
                return False

        # 内部リストにアイテム辞書を追加
        item = {
            "data": item_data,
            "base_name_var": ctk.StringVar(value=base_name),
            "original_name": original_name
        }
        self.staged_items.append(item)

        # --- UIにアイテム行を追加 ---
        row_frame = ctk.CTkFrame(self.staged_list_frame, fg_color="gray30")

        # 元ファイル名ラベル (上部)
        ctk.CTkLabel(
            row_frame, text=original_name, font=("", 10), text_color="gray70"
        ).pack(side="top", fill="x", padx=10, pady=(3, 0))

        # 編集可能なファイル名エントリー (中央)
        name_entry = ctk.CTkEntry(
            row_frame, textvariable=item["base_name_var"]
        )

        if self.merge_files_checkbox.get():
            integrated_name = ""
            # このアイテムが1番目か？ (staged_itemsに追加されたばかりなので、len==1)
            if len(self.staged_items) == 1:
                # 1番目のアイテム
                if not re.match(r"^\d{14}_", base_name):
                    now = datetime.datetime.now()
                    integrated_name = f"{now.strftime('%Y%m%d_%H%M%S')}_統合クリップ"
                else:
                    integrated_name = f"{base_name}_統合クリップ"

                item["base_name_var"].set(integrated_name)
                # 1番目のEntryは編集可能にし、同期イベントをバインド
                name_entry.configure(state="normal")
                name_entry.bind("<KeyRelease>", self.sync_merge_name)
            else:
                # 2個目以降のアイテム
                integrated_name = self.staged_items[0]["base_name_var"].get()
                item["base_name_var"].set(integrated_name)
                # 2番目以降のEntryは無効
                name_entry.configure(state="disabled")

        name_entry.pack(
            side="left", fill="x", expand=True, padx=10, pady=(0, 5))

        # 削除ボタン (右側)
        delete_btn = ctk.CTkButton(
            row_frame, text="x", width=28, height=28,
            command=lambda i=item, rf=row_frame: self.remove_staged_item(i, rf)
        )
        delete_btn.pack(side="right", padx=(0, 10), pady=(0, 5))

        row_frame.pack(fill="x", pady=5, padx=5)
        return True

    def remove_staged_item(
            self,
            item_to_remove: dict,
            row_frame: ctk.CTkFrame
    ) -> None:
        """
        指定されたアイテムを内部リストとUIから削除します。
        """
        try:
            # 削除する前に、それが1番目のアイテムだったか確認
            is_first_item = (
                len(self.staged_items) > 0 and
                self.staged_items[0] == item_to_remove
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
                if (p.is_file() and
                        p.suffix.lower() in SUPPORTED_EXTENSIONS):
                    if self.add_staged_item(p, p.stem, p.name):
                        added_files_count += 1
                elif p.is_dir():
                    for ext in SUPPORTED_EXTENSIONS:
                        for child_file in p.glob(f"**/*{ext}"):
                            if self.add_staged_item(
                                    child_file,
                                    child_file.stem,
                                    child_file.name
                                    ):
                                added_files_count += 1

            self.update_staged_files_label()
            if added_files_count > 0:
                self.parent_app.status_label.configure(
                    text=f"{added_files_count}件のファイルを追加 (D&D)")

        except Exception as e:
            messagebox.showerror(
                "ドロップエラー", f"ファイルの処理に失敗しました: {e}", parent=self)

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
                    "クリップボードの画像 " +
                    f"({pil_image.width}x{pil_image.height})"
                    )

                self.add_staged_item(pil_image, base_name, original_name)
                self.update_staged_files_label()
                self.parent_app.status_label.configure(
                    text=f"クリップボードの画像を追加 ( {base_name} )")

            else:
                # 2. 画像でない場合、テキスト (ファイルパス) として取得を試みる
                clipboard_content = self.clipboard_get()
                file_paths = [
                    line.strip().strip('"')
                    for line in re.split(r'[\n\t]', clipboard_content)
                    if line.strip()
                ]

                added_files_count = 0
                for f in file_paths:
                    p = Path(f)
                    if (p.is_file() and
                            p.suffix.lower() in SUPPORTED_EXTENSIONS):
                        if self.add_staged_item(p, p.stem, p.name):
                            added_files_count += 1
                    else:
                        logger.warning(
                            f"スキップ (無効なパスまたは非対応拡張子): {f}",
                            extra={'sensitive': True}
                        )

                if added_files_count > 0:
                    self.update_staged_files_label()
                    self.parent_app.status_label.configure(
                        text=f"{added_files_count}件のファイルパスを追加 (貼付)")

        except Exception as e:
            if "CLIPBOARD" in str(e):
                self.parent_app.status_label.configure(
                    text="クリップボードが空か、対応形式ではありません。", text_color="orange")
            else:
                messagebox.showerror(
                    "貼付エラー", f"クリップボードの処理に失敗しました: {e}", parent=self)

    def update_staged_files_label(self) -> None:
        """
        処理対象リストの件数をラベルに反映し、
        実行ボタンの有効/無効状態を切り替えます。
        """
        count = len(self.staged_items)
        self.staged_files_label.configure(text=f"処理対象ファイル: {count} 件")

        if (count > 0 and self.parent_app.font_path and
                Path(self.parent_app.font_path).is_file()):
            self.staged_run_button.configure(state="normal")
        else:
            self.staged_run_button.configure(state="disabled")

    def clear_staged_list_ui(self) -> None:
        """ (未使用) UI上のリスト表示をすべてクリアします。 """
        for widget in self.staged_list_frame.winfo_children():
            widget.destroy()

    def run_staged_process(self) -> None:
        """
        「処理実行」ボタンの処理。
        (D&Dウィンドウ専用のパイプライン)

        「統合」チェックボックスの状態に応じて、処理を分岐する。
        - OFF: ファイルごとに正規化・メタデータ付与を行う。
        - ON:  全ファイルを正規化・OCRした後、1つのPDFに連結し、最後にメタデータ付与を行う。
        """
        if not self.staged_items:
            messagebox.showinfo("情報", "処理対象のファイルが指定されていません。", parent=self)
            return

        # 1. 親アプリから設定情報を取得
        font_path = self.parent_app.font_path
        if (not font_path or not Path(font_path).is_file()):
            self.parent_app.status_label.configure(
                text="エラー: config.iniで有効なフォントパスが指定されていません。",
                text_color="orange"
            )
            return

        config_data = self.parent_app.config_data
        key_rect_tuple = config_data.get('key_rect', (0, 0, 0, 0))
        paper_width = self.parent_app.paper_width
        paper_height = self.parent_app.paper_height
        enable_tesseract = config_data.get('enable_tesseract_ocr', False)
        flatten_ink = config_data.get('flatten_ink_annotations', True)
        paper_size_str = self.parent_app.config_data.get('paper_size', 'A4')

        # 2. 出力先フォルダを選択
        dest_folder = filedialog.askdirectory(
            title="出力先フォルダを選択してください", parent=self)
        if not dest_folder:
            return
        dest_path = Path(dest_folder)
        source_folders = {
            item['data'].parent
            for item in self.staged_items
            if isinstance(item['data'], Path)
        }
        if dest_path in source_folders:
            messagebox.showerror(
                "エラー", "出力先は入力元と異なるフォルダを選択してください。", parent=self)
            return

        # 3. メタデータをUIから取得
        index_key_raw = self.index_key_combo.get()
        index_key_to_embed = ""
        text_color = None
        if index_key_raw != "（未選択）":
            index_key_to_embed = index_key_raw
            key_colors_dict = config_data.get('key_colors', {})
            hex_color = key_colors_dict.get(index_key_raw.lower())
            if hex_color:
                text_color = hex_to_rgb_tuple(hex_color)
        comment_to_embed = self.comment_textbox.get("1.0", "end-1c").strip()

        # 引用Keyリストを取得 ([[Key: Title]]形式に対応)
        cited_keys_str = self.cited_keys_entry.get("1.0", "end-1c").strip()
        cited_keys_list = []

        refs_qr_size_pt = self.parent_app.config_data.get(
            'refs_qr_size', 75  # configからQRサイズを取得
        )

        # Key (14桁以上の数字) を抽出するための正規表現
        key_regex = re.compile(r'(\d{14,})')

        if cited_keys_str:
            # 1. カンマで分割 (複数のKey/リンクが入力された場合に対応)
            parts = re.split(r'[,\n]+', cited_keys_str)

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

        sist_view_date = datetime.datetime.now().strftime('%Y-%m-%d')
        sist_string_formal = None
        sist_string_readable = None

        # 何か1つでも入力がある場合のみ、書誌情報文字列を構築
        if (raw_sist_author or raw_sist_title or
                raw_sist_site or raw_sist_date):

            sist_author = raw_sist_author or "（著者不明）"
            sist_title = raw_sist_title or "（タイトル不明）"
            sist_site = raw_sist_site or "（サイト名不明）"
            sist_date = raw_sist_date or "（更新日不明）"

            # D&D/ペーストの場合、URLは不明なため挿入しない
            sist_string_formal = (
                f"{sist_author} . “{sist_title}” . {sist_site} . "
                f"{sist_date} . (参照 {sist_view_date})"
            )
            sist_string_readable = (
                f"著者:{sist_author}\n\nページ名:\n“{sist_title}”\n\n"
                f"サイト名/出典:\n{sist_site}\n\n"
                f"更新日:{sist_date} (参照:{sist_view_date})"
            )

        # 4. 処理対象リストを作成
        items_to_process = []

        # ファイル名に使用できない文字を置換する正規表現
        invalid_chars_pattern = re.compile(r'[\\/:\*\?"<>\|]')

        for item in self.staged_items:
            # 元のファイル名を取得
            base_name_raw = item["base_name_var"].get().strip()

            # サニタイズ処理を実行
            base_name = invalid_chars_pattern.sub('_', base_name_raw)

            if not base_name:
                messagebox.showerror(
                    "ファイル名エラー", f"ファイル名が空です (元の名前: {item['original_name']})",
                    parent=self)
                return

            # もし置換が発生したら、UIの表示 (StringVar) にも反映する
            if base_name != base_name_raw:
                item["base_name_var"].set(base_name)

            items_to_process.append(
                (item['data'], base_name, type(item['data']))
            )

        # 5. 「統合」チェックボックスの状態で処理を分岐
        is_merge_mode = self.merge_files_checkbox.get()
        temp_dir = None

        try:
            temp_dir = Path(
                tempfile.mkdtemp(prefix="synapsen_dnd_", dir=dest_path)
            )

            if not is_merge_mode:
                # --- [分岐 A: 個別処理] ---
                self.run_pipeline_individual(
                    items_to_process, dest_path, temp_dir,
                    font_path, paper_width, paper_height,
                    enable_tesseract, paper_size_str,
                    key_rect_tuple, index_key_to_embed, text_color,
                    comment_to_embed,
                    sist_string_formal, sist_string_readable,
                    cited_keys_list,
                    refs_qr_size_pt,
                    flatten_ink
                )
            else:
                # --- [分岐 B: 統合処理] ---
                self.run_pipeline_merge(
                    items_to_process, dest_path, temp_dir,
                    font_path, paper_width, paper_height,
                    enable_tesseract, paper_size_str,
                    key_rect_tuple, index_key_to_embed,
                    text_color, comment_to_embed,
                    sist_string_formal, sist_string_readable,
                    cited_keys_list,
                    refs_qr_size_pt,
                    flatten_ink
                )

            # 成功したら、このウィンドウを閉じる
            self.on_close()

        except Exception as e:
            messagebox.showerror("処理エラー", f"処理中にエラーが発生しました:\n{e}", parent=self)
            self.parent_app.status_label.configure(
                text=f"エラーが発生しました: {e}")

        finally:
            if temp_dir and temp_dir.exists():
                try:
                    shutil.rmtree(temp_dir)
                except Exception as e:
                    logger.error(f"警告: 一時フォルダの削除に失敗しました: {e}")

    def run_pipeline_individual(
            self, items_to_process, dest_path, temp_dir,
            font_path, paper_width, paper_height,
            enable_tesseract, paper_size_str,
            key_rect_tuple, index_key_to_embed, text_color, comment_to_embed,
            sist_string_formal, sist_string_readable,
            cited_keys_list,
            refs_qr_size_pt,
            flatten_ink=True
            ):
        """
        個別ファイル処理パイプライン
        """
        total_files = len(items_to_process)
        for i, item_tuple in enumerate(items_to_process):
            item_data, base_name, original_type = item_tuple

            status_prefix = f"処理中 ({i+1}/{total_files}):"
            self.parent_app.status_label.configure(
                text=f"{status_prefix} {base_name}"
            )
            self.parent_app.update_idletasks()

            final_output_pdf = dest_path / f"{base_name}.pdf"
            temp_converted_pdf = temp_dir / f"conv_{base_name}.pdf"
            temp_flattened_pdf = temp_dir / f"flat_{base_name}.pdf"

            path_to_flatten: Path

            # --- 1-A: MD -> PDF ---
            if (
                    isinstance(item_data, Path) and
                    item_data.suffix.lower() == ".md"
            ):
                self.parent_app.status_label.configure(
                    text=f"{status_prefix} MD->PDF変換: {base_name}"
                )
                self.parent_app.update_idletasks()
                try:
                    temp_md_pdf = temp_dir / f"md_{base_name}.pdf"
                    convert_document_to_pdf(
                        item_data, temp_md_pdf, paper_size_str
                    )
                    item_data = temp_md_pdf  # 次の入力として上書き
                except Exception as e:
                    logger.warning(f"警告: {base_name} のMarkdown変換に失敗: {e}")
                    messagebox.showerror(
                        "Markdown変換エラー", f"{base_name} の変換に失敗しました:\n{e}",
                        parent=self
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
                text=f"{status_prefix} フラット化中: {base_name}")
            self.parent_app.update_idletasks()
            high_fidelity_flatten(
                str(path_to_flatten),
                str(temp_flattened_pdf),
                font_path,
                flatten_ink=flatten_ink
            )

            # --- 3: 正規化 ---
            self.parent_app.status_label.configure(
                text=f"{status_prefix} 正規化中: {base_name}")
            self.parent_app.update_idletasks()
            normalize_pdf_to_papersize(
                str(temp_flattened_pdf), str(final_output_pdf),
                paper_width, paper_height
            )

            # --- 4: OCR ---
            self.parent_app.status_label.configure(
                text=f"{status_prefix} OCR埋込処理中: {base_name}")
            self.parent_app.update_idletasks()
            embed_ocr_text_in_pdf(
                str(final_output_pdf), enable_tesseract,
                font_path, 'jpn+jpn_vert'
            )

            # --- 5: メタデータ追記 ---
            self.parent_app.status_label.configure(
                text=f"{status_prefix} メタデータを追記中: {base_name}")
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
                refs_qr_size_pt=refs_qr_size_pt
            )

        messagebox.showinfo(
            "完了", f"{total_files}個のファイルにメタデータを埋め込み、処理が完了しました。", parent=self
        )
        self.parent_app.status_label.configure(text="処理が完了しました。")

    def run_pipeline_merge(
            self, items_to_process, dest_path, temp_dir,
            font_path, paper_width, paper_height,
            enable_tesseract, paper_size_str,
            key_rect_tuple, index_key_to_embed, text_color, comment_to_embed,
            sist_string_formal, sist_string_readable,
            cited_keys_list,
            refs_qr_size_pt,
            flatten_ink=True
            ):
        """
        ファイル統合処理パイプライン
        """

        # 統合後のファイル名は、リストの最初のアイテムから取得
        if not items_to_process:
            return
        merged_base_name = items_to_process[0][1].strip()
        if not merged_base_name:
            messagebox.showerror("ファイル名エラー", "統合後のファイル名が空です。", parent=self)
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
            if (
                    isinstance(item_data, Path) and
                    item_data.suffix.lower() == ".md"
            ):
                self.parent_app.status_label.configure(
                    text=f"{status_prefix} MD->PDF変換: {item_original_name}"
                )
                self.parent_app.update_idletasks()
                try:
                    temp_md_pdf = temp_dir / f"md_{i}_{base_name}.pdf"
                    convert_document_to_pdf(
                        item_data, temp_md_pdf, paper_size_str
                    )
                    item_data = temp_md_pdf
                except Exception as e:
                    logger.error(
                        f"警告: {item_original_name} のMarkdown変換に失敗: {e}")
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
                text=f"{status_prefix} フラット化中: {item_original_name}")
            self.parent_app.update_idletasks()
            high_fidelity_flatten(
                str(path_to_flatten),
                str(temp_flattened_pdf),
                font_path,
                flatten_ink=flatten_ink
            )

            # --- 3: 正規化 (出力先を normalized_part_pdf に) ---
            self.parent_app.status_label.configure(
                text=f"{status_prefix} 正規化中: {item_original_name}")
            self.parent_app.update_idletasks()
            normalize_pdf_to_papersize(
                str(temp_flattened_pdf), str(normalized_part_pdf),
                paper_width, paper_height
            )

            # --- 4: OCR (normalized_part_pdf に対して実行) ---
            self.parent_app.status_label.configure(
                text=f"{status_prefix} OCR埋込処理中: {item_original_name}")
            self.parent_app.update_idletasks()
            embed_ocr_text_in_pdf(
                str(normalized_part_pdf), enable_tesseract,
                font_path, 'jpn+jpn_vert'
            )

            # 連結リストに追加
            normalized_pdf_paths.append(normalized_part_pdf)

        # --- ループ終了後 ---

        if not normalized_pdf_paths:
            messagebox.showerror("エラー", "統合できるファイルがありませんでした。", parent=self)
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
            refs_qr_size_pt=refs_qr_size_pt
        )

        messagebox.showinfo(
            "完了", f"{total_files}個のファイルを1つのPDFに統合し、処理が完了しました。", parent=self
        )
        self.parent_app.status_label.configure(
            text="処理が完了しました。", text_color="gray")

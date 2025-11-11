import customtkinter as ctk
from tkinter import filedialog, messagebox
from pathlib import Path
import datetime
import re
import tkinterdnd2
from PIL import ImageGrab
import tempfile
import shutil

from pdf_utils import (
    # D&D/画像クリップ用のメタデータ挿入関数
    add_metadata_to_image_clip,
    add_metadata_to_web_clip,
    hex_to_rgb_tuple,
    # D&Dパイプラインで個別に実行するためインポート
    convert_image_to_pdf,
    convert_pil_image_to_pdf,
    convert_markdown_to_pdf,
    high_fidelity_flatten,
    normalize_pdf_to_papersize,
    embed_ocr_text_in_pdf
)

SUPPORTED_EXTENSIONS = [
    ".pdf",
    ".png",
    ".jpg",
    ".jpeg",
    ".bmp",
    ".tiff",
    ".md"
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
        self.geometry("450x700")

        if self.parent_app.icon_path:
            try:
                self.iconbitmap(default=str(self.parent_app.icon_path))
            except Exception as e:
                print(f"Icon set error (DND Window): {e}")

        self.protocol("WM_DELETE_WINDOW", self.on_close)
        self.transient(parent_app)
        self.grab_set()

        # --- [UI定義] ---

        # 1. ドラッグ＆ドロップ ゾーン
        self.drop_target_frame = ctk.CTkFrame(
            self, height=100, fg_color="gray25")
        self.drop_target_frame.pack(pady=10, padx=10, fill="x", expand=False)

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

        # 2. メタデータ入力フレーム
        meta_frame = ctk.CTkFrame(self)
        meta_frame.pack(pady=5, padx=10, fill="x")

        # 2a. IndexKey 選択
        ctk.CTkLabel(
            meta_frame, text="IndexKey (PDF 1ページ目に埋込):", anchor="w"
            ).pack(pady=(5, 0), padx=10, fill="x")

        # 親アプリの config_data からオプションを取得
        key_options = self.parent_app.config_data.get(
                'commonplace_keys_options', [])

        self.index_key_combo = ctk.CTkComboBox(
            meta_frame,
            values=["（未選択）"] + key_options
        )
        self.index_key_combo.set("（未選択）")
        self.index_key_combo.pack(pady=5, padx=10, fill="x")

        # 2b. コメント入力
        ctk.CTkLabel(
            meta_frame, text="コメント (PDF 2ページ目に埋込):", anchor="w"
            ).pack(pady=(5, 0), padx=10, fill="x")
        self.comment_textbox = ctk.CTkTextbox(meta_frame, height=80)
        self.comment_textbox.pack(
            pady=5, padx=10, fill="both", expand=True)

        # 3. 処理対象リスト (スクロールフレーム)
        self.staged_list_frame = ctk.CTkScrollableFrame(
            self, label_text="処理対象リスト (ファイル名を編集可能)")
        self.staged_list_frame.pack(pady=5, padx=10, fill="both", expand=True)

        # 4. 処理対象ファイル数のラベル
        self.staged_files_label = ctk.CTkLabel(
            self, text="処理対象ファイル: 0 件")
        self.staged_files_label.pack(pady=5, padx=10)

        # 5. 実行ボタン
        self.staged_run_button = ctk.CTkButton(
            self,
            text="出力先を選んで処理実行",
            command=self.run_staged_process,
            state="disabled"
        )
        self.staged_run_button.pack(pady=10, padx=10, fill="x", ipady=10)

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
            print(f"DND window close error: {e}")

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
                print(f"スキップ (既に追加済み): {original_name}")
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

        Args:
            item_to_remove (dict): self.staged_items 内の削除対象アイテム辞書。
            row_frame (ctk.CTkFrame): 削除対象アイテムに対応するUIフレーム。
        """
        try:
            self.staged_items.remove(item_to_remove)
            row_frame.destroy()
            self.update_staged_files_label()
        except ValueError:
            print("エラー: 削除対象のアイテムがリストに見つかりません。")

    def handle_drop(self, event: tkinterdnd2.TkinterDnD.DnDEvent) -> None:
        """
        D&Dイベント（ファイル/フォルダドロップ）を処理します。

        Args:
            event (tkinterdnd2.TkinterDnD.DnDEvent): D&Dイベントオブジェクト。
        """
        try:
            # event.data は { } で囲まれたパス文字列のリスト (Tcl/Tk形式)
            file_paths = self.tk.splitlist(event.data)
            added_files_count = 0

            for f in file_paths:
                p = Path(f)
                if (p.is_file() and
                        p.suffix.lower() in SUPPORTED_EXTENSIONS):
                    # ファイルの場合
                    if self.add_staged_item(p, p.stem, p.name):
                        added_files_count += 1
                elif p.is_dir():
                    # フォルダの場合、再帰的に検索
                    print(f"フォルダを検索中: {p}")
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
                # 改行またはタブで区切られたパスに対応
                file_paths = [
                    line.strip().strip('"')  # 引用符も除去
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
                        print(f"スキップ (無効なパスまたは非対応拡張子): {f}")

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

        # フォントパスのチェックを親アプリ経由で行う
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

        メタデータ付与ステップ(5)で、入力タイプに応じて分岐するよう変更。
        - 画像/PIL: 1ページ目にKey, 2ページ目にComment (add_metadata_to_image_clip)
        - MD/PDF: 1ページ目にKey, 最終ページにComment (add_metadata_to_web_clip)
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

        # Pandocが必要とする設定値を取得 (Markdown連携用)
        paper_size_str = self.parent_app.config_data.get('paper_size', 'A4')

        # 2. 出力先フォルダを選択
        dest_folder = filedialog.askdirectory(
            title="出力先フォルダを選択してください", parent=self)
        if not dest_folder:
            return

        dest_path = Path(dest_folder)

        # 入力元と出力先が同一フォルダでないかチェック
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
        text_color = None  # fitzデフォルト (黒)

        if index_key_raw != "（未選択）":
            index_key_to_embed = index_key_raw
            key_colors_dict = config_data.get('key_colors', {})
            hex_color = key_colors_dict.get(index_key_raw.lower())
            if hex_color:
                text_color = hex_to_rgb_tuple(hex_color)

        comment_to_embed = self.comment_textbox.get("1.0", "end-1c").strip()

        # 4. 処理対象リストを作成
        items_to_process = []
        for item in self.staged_items:
            base_name = item["base_name_var"].get().strip()
            if not base_name:
                messagebox.showerror(
                    "ファイル名エラー", f"ファイル名が空です (元の名前: {item['original_name']})",
                    parent=self)
                return
            # (item['data'], base_name) だけでなく、
            # 元のデータ型 (Path or PIL) も渡すようにタプルを変更
            items_to_process.append(
                (item['data'], base_name, type(item['data']))
            )

        # 5. 専用パイプラインの実行
        temp_dir = None
        total_files = len(items_to_process)
        try:
            # 一時フォルダを作成
            temp_dir = Path(
                tempfile.mkdtemp(prefix="synapsen_dnd_", dir=dest_path)
            )

            for i, item_tuple in enumerate(items_to_process):
                item_data, base_name, original_type = item_tuple  # タプルの展開

                status_prefix = f"処理中 ({i+1}/{total_files}):"
                self.parent_app.status_label.configure(
                    text=f"{status_prefix} {base_name}"
                )
                self.parent_app.update_idletasks()

                # 最終的な出力パス
                final_output_pdf = dest_path / f"{base_name}.pdf"
                # 一時ファイルパス
                temp_converted_pdf = temp_dir / f"conv_{base_name}.pdf"
                temp_flattened_pdf = temp_dir / f"flat_{base_name}.pdf"

                path_to_flatten: Path  # フラット化対象のPDFパス

                # MDファイルが処理対象だったかどうかのフラグ
                is_markdown_source = False

                # --- パイプライン 1-A: MD -> PDF変換 ---
                if (
                        isinstance(item_data, Path) and
                        item_data.suffix.lower() == ".md"
                ):
                    is_markdown_source = True  # フラグを立てる
                    self.parent_app.status_label.configure(
                        text=f"{status_prefix} MD->PDF変換: {base_name}"
                    )
                    self.parent_app.update_idletasks()
                    try:
                        temp_md_pdf = temp_dir / f"md_{base_name}.pdf"

                        convert_markdown_to_pdf(
                            item_data,
                            temp_md_pdf,
                            paper_size_str
                        )
                        item_data = temp_md_pdf
                    except Exception as e:
                        print(f"警告: {base_name} のMarkdown変換に失敗: {e}")
                        messagebox.showerror(
                            "Markdown変換エラー",
                            f"{base_name} の変換に失敗しました:\n{e}",
                            parent=self
                        )
                        continue  # このファイルはスキップ

                # --- パイプライン 1-B: 画像 -> PDF変換 ---
                if isinstance(item_data, Path):
                    input_file_path = item_data
                    if input_file_path.suffix.lower() != ".pdf":
                        # 画像
                        convert_image_to_pdf(
                            input_file_path, temp_converted_pdf)
                        path_to_flatten = temp_converted_pdf
                    else:
                        # PDF (D&DされたPDF、またはMDから変換されたPDF)
                        path_to_flatten = input_file_path
                else:
                    # PIL (ペースト)
                    convert_pil_image_to_pdf(item_data, temp_converted_pdf)
                    path_to_flatten = temp_converted_pdf

                # --- パイプライン 2: フラット化 (フォームのテキスト化) ---
                high_fidelity_flatten(
                    str(path_to_flatten),
                    str(temp_flattened_pdf),
                    font_path
                )

                # --- パイプライン 3: 正規化 (サイズ統一) ---
                normalize_pdf_to_papersize(
                    str(temp_flattened_pdf),
                    str(final_output_pdf),
                    paper_width,
                    paper_height
                )

                # --- パイプライン 4: OCR埋め込み ---
                embed_ocr_text_in_pdf(
                    str(final_output_pdf),
                    enable_tesseract,
                    font_path,
                    'jpn+jpn_vert'
                )

                # --- パイプライン 5: メタデータ追記 (分岐) ---

                # original_type が Path かつ is_markdown_source が True -> MDクリップ
                # original_type が Path で、is_markdown_source が False -> PDFのD&D
                # original_type が PIL (Image) -> 画像/ペーストクリップ

                if (
                        is_markdown_source or
                        (
                            original_type == Path and
                            item_data.suffix.lower() == ".pdf"
                        )
                ):
                    # [分岐1] MD または PDF の D&D の場合
                    # 1ページ目にKey、最終ページにComment (Webクリップ方式)
                    # (書誌情報は None を渡す)
                    add_metadata_to_web_clip(
                        str(final_output_pdf),
                        font_path,
                        paper_width,
                        paper_height,
                        key_rect_tuple,
                        index_key_to_embed,
                        text_color,
                        comment_to_embed,
                        None,  # sist_string_formal
                        None   # sist_string_readable
                    )
                else:
                    # [分岐2] 画像 または PIL (ペースト) の場合
                    # 1ページ目にKey、2ページ目にComment (画像クリップ方式)
                    add_metadata_to_image_clip(
                        str(final_output_pdf),
                        font_path,
                        paper_width,
                        paper_height,
                        key_rect_tuple,
                        index_key_to_embed,
                        text_color,
                        comment_to_embed
                    )

            messagebox.showinfo(
                "完了",
                f"{total_files}個のファイルにメタデータを埋め込み、処理が完了しました。",
                parent=self
            )
            # 成功したら、このウィンドウを閉じる
            self.on_close()

        except Exception as e:
            messagebox.showerror("処理エラー", f"処理中にエラーが発生しました:\n{e}", parent=self)
            self.parent_app.status_label.configure(
                text=f"エラーが発生しました: {e}")

        finally:
            # 正常終了・異常終了に関わらず、一時フォルダを削除
            if temp_dir and temp_dir.exists():
                try:
                    shutil.rmtree(temp_dir)
                except Exception as e:
                    print(f"警告: 一時フォルダの削除に失敗しました: {e}")

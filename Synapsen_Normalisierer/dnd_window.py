import customtkinter as ctk
from tkinter import filedialog, messagebox
from pathlib import Path
import datetime
import re
import tkinterdnd2
from PIL import ImageGrab

SUPPORTED_EXTENSIONS = [".pdf", ".png", ".jpg", ".jpeg", ".bmp", ".tiff"]


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
        self.geometry("450x550")

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
                "ここにファイル (PDF, JPG, PNG) を\nドラッグ＆ドロップしてください\n\n" +
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

        # 2. 処理対象リスト (スクロールフレーム)
        self.staged_list_frame = ctk.CTkScrollableFrame(
            self, label_text="処理対象リスト (ファイル名を編集可能)")
        self.staged_list_frame.pack(pady=5, padx=10, fill="both", expand=True)

        # 3. 処理対象ファイル数のラベル
        self.staged_files_label = ctk.CTkLabel(
            self, text="処理対象ファイル: 0 件")
        self.staged_files_label.pack(pady=5, padx=10)

        # 4. 実行ボタン
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
        出力先フォルダを選択させ、親アプリの共通処理関数を呼び出します。
        """
        if not self.staged_items:
            messagebox.showinfo("情報", "処理対象のファイルが指定されていません。", parent=self)
            return

        # フォントパスの再検証
        if (not self.parent_app.font_path or
                not Path(self.parent_app.font_path).is_file()):
            self.parent_app.status_label.configure(
                text="エラー: config.iniで有効なフォントパスが指定されていません。",
                text_color="orange"
            )
            return

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

        # UIのStringVarから最新のファイル名を取得してリストを作成
        items_to_process = []
        for item in self.staged_items:
            base_name = item["base_name_var"].get().strip()
            if not base_name:
                messagebox.showerror(
                    "ファイル名エラー", f"ファイル名が空です (元の名前: {item['original_name']})",
                    parent=self)
                return
            items_to_process.append((item['data'], base_name))

        # 親アプリの共通処理関数 (execute_normalization_process) を呼び出す
        try:
            self.parent_app.execute_normalization_process(
                items_to_process, dest_path)

            # 成功したら、このウィンドウを閉じる
            self.on_close()

        except Exception as e:
            messagebox.showerror("処理エラー", f"処理中にエラーが発生しました:\n{e}", parent=self)

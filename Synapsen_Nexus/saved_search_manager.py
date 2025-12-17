import customtkinter as ctk
from tkinter import messagebox
import json
import sys
from pathlib import Path

import logging

logger = logging.getLogger(__name__)

# === 2. プロジェクトルートをパスに追加 ===
current_dir = Path(__file__).parent
root_dir = current_dir.parent
if str(root_dir) not in sys.path:
    sys.path.append(str(root_dir))

from theme import SemanticColors as Colors  # noqa: E402


# ==============================================================================
# 検索名入力用ダイアログ (アイコン適用のため独自実装)
# ==============================================================================
class SaveSearchDialog(ctk.CTkToplevel):
    def __init__(self, parent_app, current_query):
        super().__init__(parent_app)
        self.title("検索を保存")
        self.geometry("400x180")
        self.configure(
            fg_color=(Colors.BACKGROUND_HOLLOW, Colors.BACKGROUND_DARK_HOLLOW)
        )
        self.result = None

        # アイコン設定
        if hasattr(parent_app, "icon_path") and parent_app.icon_path:
            try:
                self.after(
                    200, lambda: self.iconbitmap(default=str(parent_app.icon_path))
                )
            except Exception:
                pass

        self.transient(parent_app)
        self.grab_set()

        ctk.CTkLabel(self, text="この検索に名前を付けてください:").pack(pady=(20, 5))

        self.entry = ctk.CTkEntry(self, width=300)
        self.entry.pack(pady=5)
        self.entry.focus_force()

        btn_frame = ctk.CTkFrame(self, fg_color="transparent")
        btn_frame.pack(pady=10)

        ctk.CTkButton(
            btn_frame,
            text="OK",
            width=100,
            command=self.on_ok,
            fg_color=Colors.UI_BASIC,
            text_color="black",
        ).pack(side="left", padx=10)
        ctk.CTkButton(
            btn_frame,
            text="キャンセル",
            width=100,
            command=self.destroy,
            fg_color=Colors.UI_CANCEL,
            hover_color=Colors.adjust_brightness(Colors.UI_CANCEL),
        ).pack(side="left", padx=10)

        self.bind("<Return>", lambda e: self.on_ok())
        self.bind("<Escape>", lambda e: self.destroy())

    def on_ok(self):
        val = self.entry.get().strip()
        if val:
            self.result = val
            self.destroy()
        else:
            self.entry.focus_force()

    def get_input(self):
        self.master.wait_window(self)
        return self.result


# ==============================================================================
# 検索削除用ウィンドウ
# ==============================================================================
class ManageSearchesWindow(ctk.CTkToplevel):
    def __init__(self, parent_app, search_manager):
        """
        parent_app: メインの Synapsen_Nexus インスタンス
        search_manager: SavedSearchManager のインスタンス
        """
        super().__init__(parent_app)
        self.parent_app = parent_app
        self.search_manager = search_manager
        self.title("保存済み検索の管理")
        self.geometry("450x500")
        self.configure(
            fg_color=(Colors.BACKGROUND_HOLLOW, Colors.BACKGROUND_DARK_HOLLOW)
        )

        # アイコン設定
        if self.parent_app.icon_path:
            try:
                self.after(
                    200, lambda: self.iconbitmap(default=str(self.parent_app.icon_path))
                )
            except Exception as e:
                logger.error(f"Icon set error (ManageSearches): {e}")

        self.transient(parent_app)
        self.grab_set()

        self.scroll_frame = ctk.CTkScrollableFrame(
            self,
            label_text="クリックして削除",
            fg_color=Colors.BACKGROUND_PANEL,
            label_fg_color=Colors.adjust_brightness(Colors.BACKGROUND_PANEL, 0.8),
        )
        self.scroll_frame.pack(fill="both", expand=True, padx=10, pady=10)

        self.populate_list()

    def populate_list(self):
        """self.search_manager.saved_searches に基づいてリストを(再)描画する"""
        for widget in self.scroll_frame.winfo_children():
            widget.destroy()

        # 参照先を search_manager に
        if not self.search_manager.saved_searches:
            ctk.CTkLabel(self.scroll_frame, text="保存済みの検索はありません。").pack(
                padx=10, pady=10
            )
            return

        sorted_search_names = sorted(self.search_manager.saved_searches.keys())

        for search_name in sorted_search_names:
            query = self.search_manager.saved_searches[search_name]
            row_frame = ctk.CTkFrame(
                self.scroll_frame,
                fg_color=(Colors.BACKGROUND_HOLLOW, Colors.BACKGROUND_DARK_HOLLOW),
            )

            delete_btn = ctk.CTkButton(
                row_frame,
                text="削除 (X)",
                width=80,
                fg_color=Colors.LABEL_DENGER,
                hover_color=Colors.adjust_brightness(Colors.LABEL_DENGER),
                command=lambda name=search_name: self.confirm_delete(name),
            )
            delete_btn.pack(side="left", padx=(7, 5), pady=7)

            label_text = f"名前: {search_name}\nクエリ: {query}"
            label = ctk.CTkLabel(
                row_frame,
                text=label_text,
                anchor="w",
                justify="left",
                text_color="black",
            )
            label.pack(side="left", fill="x", expand=True, padx=5, pady=7)
            row_frame.pack(fill="x", padx=5, pady=5)

    def confirm_delete(self, search_name_to_delete):
        """削除の最終確認"""
        answer = messagebox.askyesno(
            "削除の確認",
            f"以下の保存済み検索を削除しますか？\n\n名前: {search_name_to_delete}",
            parent=self,
        )

        if answer:
            try:
                # 1. マネージャーの辞書から削除
                if search_name_to_delete in self.search_manager.saved_searches:
                    del self.search_manager.saved_searches[search_name_to_delete]
                else:
                    messagebox.showerror(
                        "エラー", "削除対象が見つかりません。", parent=self
                    )
                    return

                # 2. JSONファイルへ保存
                if self.search_manager._save_searches_to_json():
                    # 3. メインアプリのComboBoxを更新
                    self.search_manager.update_saved_search_combo()
                    # 4. このウィンドウのリストを再描画
                    self.populate_list()
            except Exception as e:
                messagebox.showerror(
                    "削除エラー", f"削除に失敗しました:\n{e}", parent=self
                )


# ==============================================================================
# 保存済み検索管理クラス
# ==============================================================================
class SavedSearchManager:
    def __init__(self, parent_app):
        """
        parent_app: メインの Synapsen_Nexus インスタンス
        """
        self.parent_app = parent_app  # メインアプリのUIを参照・操作するために保持
        self.saved_searches = {}
        self.saved_searches_path = None

    def load_saved_searches(self, base_path):
        """
        saved_searches.json ファイルから保存済み検索を読み込み、
        UI (ComboBox) に反映する。
        """
        # base_path を main から受け取る
        self.saved_searches_path = base_path / "saved_searches.json"

        if self.saved_searches_path and self.saved_searches_path.is_file():
            try:
                with open(self.saved_searches_path, "r", encoding="utf-8") as f:
                    self.saved_searches = json.load(f)
            except Exception as e:
                logger.error(f"saved_searches.json の読み込みに失敗: {e}")
                self.saved_searches = {}
        else:
            self.saved_searches = {}

        # UIのComboBoxを更新
        self.update_saved_search_combo()

    def _save_searches_to_json(self):
        """
        現在の self.saved_searches 辞書を json ファイルに上書き保存する。
        """
        if not self.saved_searches_path:
            messagebox.showerror(
                "エラー",
                "保存先パス(saved_searches.json)が設定されていません。",
                parent=self.parent_app,
            )
            return False
        try:
            with open(self.saved_searches_path, "w", encoding="utf-8") as f:
                json.dump(self.saved_searches, f, indent=4, ensure_ascii=False)
            return True
        except Exception as e:
            messagebox.showerror(
                "保存エラー",
                f"検索内容の保存に失敗しました:\n{e}",
                parent=self.parent_app,
            )
            return False

    def open_manage_searches_window(self):
        """
        「管理」ボタン押下時。保存済み検索の管理ウィンドウを開く。
        """
        # self (マネージャー自身) も渡す
        manage_win = ManageSearchesWindow(self.parent_app, self)
        manage_win.focus()

    def update_saved_search_combo(self):
        """
        self.saved_searches の内容に基づき、メインアプリのComboBoxの選択肢を更新する。
        """
        # 起動直後などUI未生成時はスキップ
        if not hasattr(self.parent_app, "saved_search_combo"):
            return

        # メインアプリのUI (parent_app) を操作
        combo_box = self.parent_app.saved_search_combo

        # 基本の選択肢リスト
        search_names = ["保存済み検索..."]

        # 管理用の特別項目を追加 (区切り線的に少し目立たせる)
        search_names.append("【 管理... 】")

        if self.saved_searches:
            # その後に検索項目を追加
            search_names.extend(sorted(list(self.saved_searches.keys())))

        combo_box.configure(values=search_names)

        # 現在の表示をリセット
        combo_box.set("保存済み検索...")

    def on_saved_search_selected(self, selected_name: str):
        """
        ComboBoxで保存済み検索が選択されたときに呼び出される。
        """
        # 何もしない選択肢
        if selected_name == "保存済み検索...":
            return

        # 管理コマンドが選ばれた場合
        if selected_name == "【 管理... 】":
            # 選択表示を元に戻してからウィンドウを開く
            self.parent_app.saved_search_combo.set("保存済み検索...")
            self.open_manage_searches_window()
            return

        # 通常の検索実行
        if selected_name in self.saved_searches:
            query = self.saved_searches[selected_name]

            # メインアプリのUIとメソッドを操作
            self.parent_app.search_entry.delete(0, "end")
            self.parent_app.search_entry.insert(0, query)
            self.parent_app.perform_search()

        self.parent_app.saved_search_combo.set("保存済み検索...")

    def save_current_search(self):
        """
        「検索保存」ボタン押下時。現在の検索クエリを保存する。
        """
        # メインアプリのUIからクエリ取得
        current_query = self.parent_app.search_entry.get().strip()
        if not current_query:
            messagebox.showwarning(
                "検索保存",
                "保存する検索クエリが入力されていません。",
                parent=self.parent_app,
            )
            return

        # カスタムダイアログを使用
        dialog = SaveSearchDialog(self.parent_app, current_query)
        search_name = dialog.get_input()

        if not search_name:
            return  # キャンセルされた

        self.saved_searches[search_name] = current_query

        if self._save_searches_to_json():
            self.update_saved_search_combo()
            messagebox.showinfo(
                "保存完了",
                f"検索「{search_name}」を保存しました。",
                parent=self.parent_app,
            )

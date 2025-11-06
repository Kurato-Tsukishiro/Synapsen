import customtkinter as ctk
from tkinter import messagebox
import json


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

        # アイコン設定
        if self.parent_app.icon_path:
            try:
                self.iconbitmap(default=str(self.parent_app.icon_path))
            except Exception as e:
                print(f"Icon set error (ManageSearches): {e}")

        self.transient(parent_app)
        self.grab_set()

        self.scroll_frame = ctk.CTkScrollableFrame(self, label_text="クリックして削除")
        self.scroll_frame.pack(fill="both", expand=True, padx=10, pady=10)

        self.populate_list()

    def populate_list(self):
        """ self.search_manager.saved_searches に基づいてリストを(再)描画する """
        for widget in self.scroll_frame.winfo_children():
            widget.destroy()

        # 参照先を search_manager に
        if not self.search_manager.saved_searches:
            ctk.CTkLabel(
                self.scroll_frame, text="保存済みの検索はありません。"
                ).pack(padx=10, pady=10)
            return

        sorted_search_names = sorted(self.search_manager.saved_searches.keys())

        for search_name in sorted_search_names:
            query = self.search_manager.saved_searches[search_name]
            row_frame = ctk.CTkFrame(self.scroll_frame, fg_color="gray25")

            delete_btn = ctk.CTkButton(
                row_frame,
                text="削除 (X)",
                width=80,
                fg_color="#D9534F",
                hover_color="#C9302C",
                command=lambda name=search_name: self.confirm_delete(name)
            )
            delete_btn.pack(side="left", padx=(7, 5), pady=7)

            label_text = f"名前: {search_name}\nクエリ: {query}"
            label = ctk.CTkLabel(
                row_frame,
                text=label_text,
                anchor="w",
                justify="left"
            )
            label.pack(side="left", fill="x", expand=True, padx=5, pady=7)
            row_frame.pack(fill="x", padx=5, pady=5)

    def confirm_delete(self, search_name_to_delete):
        """ 削除の最終確認 """
        answer = messagebox.askyesno(
            "削除の確認",
            f"以下の保存済み検索を削除しますか？\n\n名前: {search_name_to_delete}",
            parent=self
        )

        if answer:
            try:
                # 1. マネージャーの辞書から削除
                if search_name_to_delete in self.search_manager.saved_searches:
                    del self.search_manager.saved_searches[
                        search_name_to_delete]
                else:
                    messagebox.showerror("エラー", "削除対象が見つかりません。", parent=self)
                    return

                # 2. JSONファイルへ保存
                if self.search_manager._save_searches_to_json():
                    # 3. メインアプリのComboBoxを更新
                    self.search_manager.update_saved_search_combo()
                    # 4. このウィンドウのリストを再描画
                    self.populate_list()
            except Exception as e:
                messagebox.showerror("削除エラー", f"削除に失敗しました:\n{e}", parent=self)


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
                with open(
                        self.saved_searches_path, 'r', encoding='utf-8') as f:
                    self.saved_searches = json.load(f)
            except Exception as e:
                print(f"saved_searches.json の読み込みに失敗: {e}")
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
                "エラー", "保存先パス(saved_searches.json)が設定されていません。",
                parent=self.parent_app)
            return False
        try:
            with open(self.saved_searches_path, 'w', encoding='utf-8') as f:
                json.dump(self.saved_searches, f, indent=4, ensure_ascii=False)
            return True
        except Exception as e:
            messagebox.showerror(
                "保存エラー", f"検索内容の保存に失敗しました:\n{e}", parent=self.parent_app)
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
        # メインアプリのUI (parent_app) を操作
        combo_box = self.parent_app.saved_search_combo

        if not self.saved_searches:
            combo_box.configure(values=["保存済み検索..."])
            combo_box.set("保存済み検索...")
        else:
            search_names =\
                ["保存済み検索..."] + sorted(list(self.saved_searches.keys()))
            combo_box.configure(values=search_names)
            combo_box.set("保存済み検索...")

    def on_saved_search_selected(self, selected_name: str):
        """
        ComboBoxで保存済み検索が選択されたときに呼び出される。
        """
        if selected_name == "保存済み検索...":
            return

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
                "検索保存", "保存する検索クエリが入力されていません。", parent=self.parent_app)
            return

        dialog = ctk.CTkInputDialog(
            text="この検索に名前を付けてください:", title="検索を保存"
        )
        if self.parent_app.icon_path:
            try:
                dialog.iconbitmap(default=str(self.parent_app.icon_path))
            except Exception:
                pass

        search_name = dialog.get_input()

        if not search_name:
            return  # キャンセルされた

        self.saved_searches[search_name] = current_query

        if self._save_searches_to_json():
            self.update_saved_search_combo()
            messagebox.showinfo(
                "保存完了", f"検索「{search_name}」を保存しました。", parent=self.parent_app)

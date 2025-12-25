import sys
from pathlib import Path
import customtkinter as ctk
import sqlite3
from utils import get_all_tags_with_count

current_dir = Path(__file__).parent
root_dir = current_dir.parent
if str(root_dir) not in sys.path:
    sys.path.append(str(root_dir))

from theme import SemanticColors as Colors  # noqa: E402


class TagWindow(ctk.CTkToplevel):
    def __init__(self, master, db_path, on_tag_click_callback):
        super().__init__(master)

        self.title("Tag List")
        self.geometry("500x650")
        self.configure(
            fg_color=(Colors.BACKGROUND_HOLLOW, Colors.BACKGROUND_DARK_HOLLOW)
        )
        self.db_path = db_path
        self.callback = on_tag_click_callback

        # データのキャッシュとソート状態
        self.tags_cache = []  # (tag_name, count) のリスト
        self.sort_mode = "count"  # "count" or "name"
        self.sort_desc = True  # True=降順(▼), False=昇順(▲)

        def set_window(can_use_icon: bool) -> None:
            """
            アイコンを設置した後、ウィンドウを設定する。

            Args:
                can_use_icon (bool): アイコンの適応が可能か。
            """
            if can_use_icon:
                self.iconbitmap(master.icon_path)

            # ウィンドウ設定
            self.transient(master)
            self.lift()
            self.focus_force()

        # アイコン・ウィンドウ設定
        if hasattr(master, "icon_path") and master.icon_path:
            try:
                # アイコンの設定が可能なら、遅延して反映する。
                self.after(200, lambda: set_window(True))
            except Exception:
                # できないなら、ウィンドウの設定のみ行う。
                set_window(False)

        # --- ヘッダー・コントロール部分 ---
        self.header_frame = ctk.CTkFrame(self, fg_color="transparent")
        self.header_frame.pack(fill="x", padx=10, pady=10)

        # タイトル
        self.label = ctk.CTkLabel(
            self.header_frame, text="登録タグ一覧", font=("sans-serif", 18, "bold")
        )
        self.label.pack(side="top", pady=(0, 5))

        # ソートボタン用フレーム
        self.sort_frame = ctk.CTkFrame(self.header_frame, fg_color="transparent")
        self.sort_frame.pack(side="top", pady=5)

        # ソートボタン: 件数
        self.btn_sort_count = ctk.CTkButton(
            self.sort_frame,
            text="件数 ▼",
            width=90,
            fg_color=Colors.UI_BASIC,
            hover_color=Colors.adjust_brightness(Colors.UI_BASIC),
            text_color="black",
            command=self.toggle_sort_count,
        )
        self.btn_sort_count.pack(side="left", padx=5)

        # ソートボタン: 名前
        self.btn_sort_name = ctk.CTkButton(
            self.sort_frame,
            text="名前",
            width=90,
            fg_color="transparent",
            border_width=1,
            hover_color=Colors.adjust_brightness(Colors.UI_BASIC),
            text_color=Colors.UI_CANCEL,
            command=self.toggle_sort_name,
        )
        self.btn_sort_name.pack(side="left", padx=5)

        # AND/OR 切り替えスイッチ
        self.mode_var = ctk.StringVar(value="AND")
        self.mode_segment = ctk.CTkSegmentedButton(
            self.header_frame,
            values=["AND", "OR"],
            variable=self.mode_var,
            width=150,
            fg_color=(
                Colors.adjust_brightness(Colors.BACKGROUND_HOLLOW),
                Colors.adjust_brightness(Colors.BACKGROUND_DARK_HOLLOW, 1.2),
            ),
            selected_color=Colors.UI_BASIC,
            selected_hover_color=Colors.adjust_brightness(Colors.UI_BASIC),
            unselected_color=(
                Colors.adjust_brightness(Colors.BACKGROUND_HOLLOW),
                Colors.adjust_brightness(Colors.BACKGROUND_DARK_HOLLOW, 1.2),
            ),
            unselected_hover_color=(
                Colors.adjust_brightness(Colors.BACKGROUND_HOLLOW, 0.6),
                Colors.adjust_brightness(Colors.BACKGROUND_DARK_HOLLOW, 1.4),
            ),
            text_color="black",
            text_color_disabled=(
                Colors.adjust_brightness(Colors.BACKGROUND_HOLLOW, 1.4),
                Colors.adjust_brightness(Colors.BACKGROUND_DARK_HOLLOW, 0.6),
            ),
        )
        self.mode_segment.pack(side="top", pady=(10, 5))

        # 説明ラベル
        self.info_label = ctk.CTkLabel(
            self.header_frame,
            text="クリックで現在の検索条件に追加します",
            font=("sans-serif", 12),
            text_color="gray",
        )
        self.info_label.pack(side="top")

        # --- タグリスト部分 ---
        self.scroll_frame = ctk.CTkScrollableFrame(
            self,
            fg_color=(Colors.BACKGROUND_PANEL, Colors.BACKGROUND_DARK_HOLLOW),
        )
        self.scroll_frame.pack(fill="both", expand=True, padx=10, pady=(0, 10))

        # 初期読み込み
        self.fetch_tags()

    def fetch_tags(self):
        """DBからタグを取得してキャッシュする"""
        try:
            conn = sqlite3.connect(f"file:{self.db_path}?mode=ro", uri=True)
            self.tags_cache = get_all_tags_with_count(conn)
            conn.close()

            # 取得後に描画
            self.render_tags()

        except Exception as e:
            ctk.CTkLabel(self.scroll_frame, text=f"エラー: {e}").pack()

    def toggle_sort_count(self):
        """件数ソート切り替え"""
        if self.sort_mode == "count":
            # 既に件数モードなら、昇順/降順を反転
            self.sort_desc = not self.sort_desc
        else:
            # 名前モードから切り替えなら、件数(降順)をデフォルトにする
            self.sort_mode = "count"
            self.sort_desc = True

        self.render_tags()

    def toggle_sort_name(self):
        """名前ソート切り替え"""
        if self.sort_mode == "name":
            # 既に名前モードなら、昇順/降順を反転
            self.sort_desc = not self.sort_desc
        else:
            # 件数モードから切り替えなら、名前(昇順:あいうえお順)をデフォルトにする
            self.sort_mode = "name"
            self.sort_desc = False

        self.render_tags()

    def update_sort_buttons_ui(self):
        """ボタンの見た目（アクティブ状態・矢印）を更新"""
        active_fg = Colors.UI_BASIC
        inactive_fg = "transparent"
        active_text = "black"
        inactive_text = Colors.UI_CANCEL

        # 矢印決定
        arrow = "▼" if self.sort_desc else "▲"

        if self.sort_mode == "count":
            # 件数ボタンをアクティブ化
            self.btn_sort_count.configure(
                text=f"件数 {arrow}",
                fg_color=active_fg,
                text_color=active_text,
                border_width=0,
            )
            # 名前ボタンを非アクティブ化
            self.btn_sort_name.configure(
                text="名前",
                fg_color=inactive_fg,
                text_color=inactive_text,
                border_width=1,
            )
        else:
            # 名前ボタンをアクティブ化
            self.btn_sort_name.configure(
                text=f"名前 {arrow}",
                fg_color=active_fg,
                text_color=active_text,
                border_width=0,
            )
            # 件数ボタンを非アクティブ化
            self.btn_sort_count.configure(
                text="件数",
                fg_color=inactive_fg,
                text_color=inactive_text,
                border_width=1,
            )

    def render_tags(self):
        """現在のソート設定に従ってタグを並べ替えて表示"""
        # 1. UI状態更新
        self.update_sort_buttons_ui()

        # 2. ソート実行

        if self.sort_mode == "count":
            # 件数メイン、同じなら名前順
            # desc(True) -> 件数多い順 (-x[1]), asc(False) -> 件数少ない順 (x[1])
            def key_func(x):
                return (-x[1], x[0]) if self.sort_desc else (x[1], x[0])

        else:
            # 名前メイン
            # desc(True) -> 逆順, asc(False) -> 五十音順
            def key_func(x):
                return x[0]

        # 名前モードの降順だけ reverse=True で対応（文字列比較のため）
        is_reverse = self.sort_mode == "name" and self.sort_desc

        sorted_tags = sorted(self.tags_cache, key=key_func, reverse=is_reverse)

        # 3. リストクリア
        for widget in self.scroll_frame.winfo_children():
            widget.destroy()

        if not sorted_tags:
            ctk.CTkLabel(self.scroll_frame, text="タグが見つかりません").pack(pady=20)
            return

        # 4. ボタン再生成
        for tag, count in sorted_tags:
            btn = ctk.CTkButton(
                self.scroll_frame,
                text=f"{tag} ({count})",
                fg_color="transparent",
                border_width=1,
                border_color="#888",
                hover_color=(
                    Colors.adjust_brightness(Colors.BACKGROUND_HOLLOW),
                    Colors.adjust_brightness(Colors.BACKGROUND_DARK_PANEL, 1.2),
                ),
                text_color=("black", "white"),
                anchor="w",
                height=30,
                command=lambda t=tag: self.on_click(t),
            )
            btn.pack(fill="x", pady=2, padx=5)

    def on_click(self, tag):
        """タグクリック時の処理"""
        current_query = ""
        if hasattr(self.master, "search_entry"):
            current_query = self.master.search_entry.get().strip()

        new_condition = f"tag:{tag}"
        operator = self.mode_var.get()  # "AND" or "OR"

        final_query = ""

        if not current_query:
            final_query = new_condition
        else:
            if new_condition in current_query:
                pass
            final_query = f"{current_query} {operator} {new_condition}"

        if self.callback:
            self.callback(final_query)

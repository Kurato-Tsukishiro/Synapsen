import re
import customtkinter as ctk

import logging

logger = logging.getLogger("Nexus.search_Mixin")


class NexusSearchMixin:
    """検索バーのオートコンプリートと入力制御を担当するMixin"""

    def setup_search_variables(self):
        """検索関連の変数を初期化"""
        self.search_timer = None
        self.suggestion_timer = None
        self._last_suggestion_args = None
        self.selected_suggestion_index = -1
        self.current_suggestions = []
        # autocomplete_frame は ui_mixin で作成されます

    def handle_keyrelease(self, event):
        """検索バーでのキー入力イベントハンドラ"""
        ignored_keys = (
            "Return",
            "Escape",
            "Left",
            "Up",
            "Down",
            "Right",
            "Home",
            "End",
            "Prior",
            "Next",
            "Control_L",
            "Control_R",
            "Shift_L",
            "Shift_R",
            "Alt_L",
            "Alt_R",
            "Tab",
            "F1",
            "F2",
            "F3",
            "F4",
            "F5",
            "F6",
            "F7",
            "F8",
            "F9",
            "F10",
            "F11",
            "F12",
        )
        if event.keysym in ignored_keys:
            return

        is_ctrl = (event.state & 0x0004) != 0
        if is_ctrl:
            return

        self.schedule_suggestions()
        self.schedule_search()

    def schedule_suggestions(self, event=None):
        """オートコンプリートのスケジュール"""
        if self.suggestion_timer:
            self.after_cancel(self.suggestion_timer)
            self.suggestion_timer = None

        query = self.search_entry.get()
        cursor_pos = self.search_entry.index(ctk.INSERT)
        query_to_cursor = query[:cursor_pos]

        # tag:プレフィックスの検出 (例: "tag:Id")
        tag_value_pattern = r"(?i)(^|\s|\()tags?:([^\s\)]*)$"
        match_value = re.search(tag_value_pattern, query_to_cursor)

        # tag:直後の検出 (例: "tag:")
        tag_prefix_pattern = r"(?i)(^|\s|\()tags?:\s*$"
        match_prefix = re.search(tag_prefix_pattern, query_to_cursor)

        if match_prefix or match_value:
            self.suggestion_timer = self.after(
                200,
                lambda q=query, c=cursor_pos, m=match_value: self.update_suggestions(
                    q, c, m
                ),
            )
        else:
            self.hide_autocomplete()
            self._last_suggestion_args = None

    def update_suggestions(self, query, cursor_pos, match_value):
        """候補リストを更新して表示"""
        self.selected_suggestion_index = -1
        last_tag_word = match_value.group(2).strip() if match_value else ""

        # メインクラスの属性を参照
        target_list = getattr(self, "all_unique_tags", []) or getattr(
            self, "predefined_tags", []
        )

        if last_tag_word == "":
            suggestions = target_list
        else:
            last_word_lower = last_tag_word.lower()
            suggestions = [
                t for t in target_list if t.lower().startswith(last_word_lower)
            ]

        self._last_suggestion_args = (query, cursor_pos, match_value)

        if suggestions:
            self.show_autocomplete(suggestions, query, cursor_pos, match_value)
        else:
            self.hide_autocomplete()

    def show_autocomplete(self, suggestions, query, cursor_pos, match_value):
        """オートコンプリートウィンドウの描画"""
        self.current_suggestions = suggestions

        # 既存ボタンの削除
        for widget in self.autocomplete_frame.winfo_children():
            widget.destroy()

        for i, suggestion in enumerate(suggestions):
            fg_color = (
                "gray30" if i == self.selected_suggestion_index else "transparent"
            )
            btn = ctk.CTkButton(
                self.autocomplete_frame,
                text=suggestion,
                fg_color=fg_color,
                anchor="w",
                text_color=ctk.ThemeManager.theme["CTkLabel"]["text_color"],
                command=lambda s=suggestion: self.select_suggestion(
                    s, query, cursor_pos, match_value
                ),
            )
            btn.pack(fill="x", padx=5, pady=2)

        # 検索バーの下に配置
        x = self.search_entry.winfo_rootx() - self.winfo_rootx()
        y = (
            self.search_entry.winfo_rooty()
            - self.winfo_rooty()
            + self.search_entry.winfo_height()
        )
        width = self.search_entry.winfo_width()
        height = min(200, len(suggestions) * 35)

        self.autocomplete_frame.configure(width=width, height=height)
        self.autocomplete_frame.place(x=x, y=y)
        self.autocomplete_frame.lift()

    def select_suggestion(self, suggestion, query, cursor_pos, match_value):
        """候補を選択して入力に反映"""
        prefix_part = ""
        suffix_part = query[cursor_pos:]

        if match_value:
            # マッチした箇所の開始位置まで
            prefix_part = query[: match_value.start(2)]
        else:
            prefix_part = query[:cursor_pos]
            if not prefix_part.endswith(" "):
                suggestion = " " + suggestion

        new_query = f"{prefix_part}{suggestion} {suffix_part}"
        new_cursor_pos = len(prefix_part) + len(suggestion) + 1

        self.search_entry.delete(0, "end")
        self.search_entry.insert(0, new_query)
        self.search_entry.focus_force()
        self.search_entry.icursor(new_cursor_pos)

        self.hide_autocomplete()
        self._trigger_search_now()

    def hide_autocomplete(self, event=None):
        """オートコンプリートを非表示"""
        self.after(200, lambda: self.autocomplete_frame.place_forget())

    def navigate_suggestions(self, event):
        """矢印キーでの候補移動"""
        if not self.autocomplete_frame.winfo_ismapped() or not self.current_suggestions:
            return

        num = len(self.current_suggestions)
        if event.keysym == "Down":
            self.selected_suggestion_index = (self.selected_suggestion_index + 1) % num
        elif event.keysym == "Up":
            self.selected_suggestion_index = (
                self.selected_suggestion_index - 1 + num
            ) % num

        # スクロール位置の調整
        self.autocomplete_frame._parent_canvas.yview_moveto(
            self.selected_suggestion_index / num
        )

        # ハイライト更新
        if self._last_suggestion_args:
            q, c, m = self._last_suggestion_args
            self.show_autocomplete(self.current_suggestions, q, c, m)

        return "break"

    def confirm_suggestion(self, event):
        """Enterキーで確定"""
        if (
            self.autocomplete_frame.winfo_ismapped()
            and self.selected_suggestion_index != -1
            and self._last_suggestion_args is not None
        ):
            q, c, m = self._last_suggestion_args
            self.select_suggestion(
                self.current_suggestions[self.selected_suggestion_index], q, c, m
            )
            return "break"

        # 候補選択中でなければ検索実行
        self._trigger_search_now()
        self.hide_autocomplete()

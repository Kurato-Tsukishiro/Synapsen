import logging

logger = logging.getLogger(__name__)


class ListNavigatorMixin:
    """
    Synapsen Nexus のリスト操作・キーボードナビゲーション機能を提供するMixinクラス。
    メインの Synapsen_Nexus クラスに継承させて使用します。
    """

    def setup_navigation_variables(self):
        """
        __init__ で呼び出す変数の初期化
        """
        self.list_item_widgets = []  # ウィジェット参照リスト
        self.list_cursor_index = -1  # 現在のカーソル位置
        self.list_anchor_index = -1  # 範囲選択の始点

    def setup_navigation_shortcuts(self):
        """
        キーボードショートカットの設定
        _setup_shortcuts 内から呼び出してください。
        """
        # 矢印キー (Up/Down)
        self.bind("<Up>", lambda e: self._handle_up_key(e, range_select=False))
        self.bind("<Down>", lambda e: self._handle_down_key(e, range_select=False))
        self.bind("<Shift-Up>", lambda e: self._handle_up_key(e, range_select=True))
        self.bind("<Shift-Down>", lambda e: self._handle_down_key(e, range_select=True))

        # Home / End
        self.bind("<Home>", lambda e: self._handle_home_key(e, range_select=False))
        self.bind("<End>", lambda e: self._handle_end_key(e, range_select=False))
        self.bind("<Shift-Home>", lambda e: self._handle_home_key(e, range_select=True))
        self.bind("<Shift-End>", lambda e: self._handle_end_key(e, range_select=True))

        # PageUp / PageDown
        self.bind("<Prior>", lambda e: self._handle_page_up_key(e, range_select=False))
        self.bind("<Next>", lambda e: self._handle_page_down_key(e, range_select=False))
        self.bind(
            "<Shift-Prior>", lambda e: self._handle_page_up_key(e, range_select=True)
        )
        self.bind(
            "<Shift-Next>", lambda e: self._handle_page_down_key(e, range_select=True)
        )

        # 決定・選択操作
        self.bind("<Return>", self._handle_enter_key)
        self.bind("<Shift-Return>", lambda e: self._handle_enter_key(e, pdf_mode=True))
        self.bind("<space>", self._handle_space_key)

    # =========================================================================
    # ヘルパーメソッド (ロジック本体)
    # =========================================================================

    def _set_list_cursor(
        self, index, update_anchor=True, do_range_select=False, keep_existing=False
    ):
        """
        カーソル移動処理
        index: 移動先のインデックス
        update_anchor: Trueならアンカー位置もカーソル位置に更新（通常移動）
        do_range_select: Trueならアンカーから現在位置までを範囲選択（Shift移動）
        """
        if not self.list_item_widgets:
            return

        # 範囲制限
        if index < 0:
            index = 0
        elif index >= len(self.list_item_widgets):
            index = len(self.list_item_widgets) - 1

        # 以前のカーソルのハイライト解除
        if 0 <= self.list_cursor_index < len(self.list_item_widgets):
            self.list_item_widgets[self.list_cursor_index]["frame"].configure(
                fg_color="transparent"
            )

        self.list_cursor_index = index

        # 新しいカーソル位置をハイライト
        curr_item = self.list_item_widgets[index]
        # ライト/ダーク対応の色 (必要に応じて調整)
        highlight_color = ("#D3D3D3", "#404040")
        curr_item["frame"].configure(fg_color=highlight_color)

        self._scroll_to_index(index)

        # 範囲選択ロジック
        if do_range_select:
            if self.list_anchor_index == -1:
                self.list_anchor_index = index
            self._select_range(
                self.list_anchor_index, index, keep_existing=keep_existing
            )
        elif update_anchor:
            self.list_anchor_index = index

    def _scroll_to_index(self, index):
        """指定されたインデックスが表示されるようにスクロールする"""
        if not self.list_item_widgets:
            return
        total = len(self.list_item_widgets)
        if total == 0:
            return

        # 表示領域の割合 (概算)
        pos = index / total
        try:
            # self.results_list はメインクラス側にある想定
            if index < 3:
                self.results_list._parent_canvas.yview_moveto(0)
            elif index > total - 3:
                self.results_list._parent_canvas.yview_moveto(1)
            else:
                self.results_list._parent_canvas.yview_moveto(pos - 0.05)
        except Exception:
            pass

    def _select_range(self, start_idx, end_idx, keep_existing=False):
        """
        始点から終点までの範囲を選択状態にする
        keep_existing: Trueなら既存の選択を解除しない（追加選択モード）
        """
        if not self.list_item_widgets:
            return

        lower = min(start_idx, end_idx)
        upper = max(start_idx, end_idx)

        # 1. 内部の選択セットをクリア (シフトキー無効時のみ)
        if not keep_existing:
            self.selected_keys.clear()  # self.selected_keys はメインクラス側

        # 2. UI上のチェックボックスと内部セットを一括更新
        for i, item in enumerate(self.list_item_widgets):
            key = item["key"]
            chk_var = item["chk_var"]

            if lower <= i <= upper:  # 範囲内: 強制的に選択 (ON)
                self.selected_keys.add(key)
                if chk_var.get() != "on":
                    chk_var.set("on")
            elif not keep_existing:  # 範囲外: 解除 (OFF)
                if chk_var.get() != "off":
                    chk_var.set("off")

        # 3. UI更新メソッド呼び出し
        if hasattr(self, "update_selection_ui_state"):
            self.update_selection_ui_state()

    def _move_cursor(self, step, range_select=False, keep_existing=False):
        """ "カーソルを相対移動させる (Up/Downキー用)"""
        if not self.list_item_widgets:
            return

        # カーソル未設定時の初期化
        if self.list_cursor_index == -1:
            start_index = 0 if step > 0 else len(self.list_item_widgets) - 1
            # 範囲選択開始ならアンカーも設定
            if range_select:
                self.list_anchor_index = start_index
            self._set_list_cursor(
                start_index,
                update_anchor=not range_select,
                do_range_select=range_select,
                keep_existing=keep_existing,
            )
            return

        new_index = self.list_cursor_index + step
        self._set_list_cursor(
            new_index,
            update_anchor=not range_select,
            do_range_select=range_select,
            keep_existing=keep_existing,
        )

    def _move_cursor_absolute(
        self, target_index, range_select=False, keep_existing=False
    ):
        """指定位置へ絶対移動"""
        if not self.list_item_widgets:
            return

        # -1 なら末尾を計算
        if target_index < 0:
            target_index = len(self.list_item_widgets) - 1

        self._set_list_cursor(
            target_index,
            update_anchor=not range_select,
            do_range_select=range_select,
            keep_existing=keep_existing,
        )

    def _move_cursor_relative(self, step, range_select=False, keep_existing=False):
        """ページ送り移動"""
        self._move_cursor(step, range_select, keep_existing)

    # =========================================================================
    # アクション・ヘルパーメソッド
    # =========================================================================

    def _is_ctrl_pressed(self, event):
        # 0x0004 は Controlキーのマスクビット
        return (event.state & 0x0004) != 0

    def _is_search_entry_focused(self):
        """検索バーにフォーカスがあるか判定"""
        try:
            # self.search_entry はメインクラス側にある想定
            return self.focus_get() == self.search_entry
        except KeyError:
            return False

    def _toggle_cursor_selection(self):
        """現在のカーソル行の選択状態(チェックボックス)を切り替える (Spaceキー)"""
        if 0 <= self.list_cursor_index < len(self.list_item_widgets):
            item = self.list_item_widgets[self.list_cursor_index]
            chk_var = item["chk_var"]
            key = item["key"]

            # 値を反転
            new_val = "on" if chk_var.get() == "off" else "off"
            chk_var.set(new_val)

            # 処理実行 (メインクラスのメソッドを呼び出す)
            if hasattr(self, "toggle_note_selection"):
                self.toggle_note_selection(key, chk_var)

    def _activate_cursor_item(self, open_pdf=False):
        """現在のカーソル行を決定する (Enterキー)"""
        if 0 <= self.list_cursor_index < len(self.list_item_widgets):
            item = self.list_item_widgets[self.list_cursor_index]
            row_data = item["data"]

            if open_pdf:
                # Shift+Enter: PDFを開く (メインクラスのメソッド)
                if hasattr(self, "open_pdf"):
                    self.open_pdf(row_data)
            else:
                # Enter: 詳細表示 (メインクラスのメソッド)
                if hasattr(self, "show_details"):
                    self.show_details(row_data)

    # =========================================================================
    # ハンドラメソッド (ロジック本体)
    # =========================================================================

    def _handle_up_key(self, event, range_select=False):
        """上矢印キーのハンドラ"""
        if self._is_search_entry_focused():
            # 検索バーにいる場合、オートコンプリート操作へ委譲
            self.navigate_suggestions(event)
        else:
            # それ以外ならリスト移動
            is_ctrl = self._is_ctrl_pressed(event)  # Ctrlキーの状態を確認して渡す
            self._move_cursor(-1, range_select=range_select, keep_existing=is_ctrl)
            return "break"

    def _handle_down_key(self, event, range_select=False):
        """下矢印キーのハンドラ"""
        if self._is_search_entry_focused():
            # 検索バーにいる場合
            if self.autocomplete_frame.winfo_ismapped():
                # オートコンプリートが出ていればそちらを操作
                self.navigate_suggestions(event)
            else:
                # オートコンプリートが出ていなければリストへフォーカス移動
                self.focus_set()  # 検索バーからフォーカスを外す
                # Ctrlキーの状態を確認して渡す
                is_ctrl = self._is_ctrl_pressed(event)

                # 現在のカーソル位置が有効なら、そこを維持（もしくは +1）する
                target_index = 0
                if self.list_cursor_index != -1:
                    target_index = self.list_cursor_index

                self._set_list_cursor(
                    target_index, do_range_select=range_select, keep_existing=is_ctrl
                )
        else:
            # リスト移動
            is_ctrl = self._is_ctrl_pressed(event)  # Ctrlキーの状態を確認して渡す
            self._move_cursor(1, range_select=range_select, keep_existing=is_ctrl)
            return "break"

    def _handle_enter_key(self, event, pdf_mode=False):
        """Enterキーのハンドラ"""
        if self._is_search_entry_focused():
            # 検索バーなら検索実行またはオートコンプリート確定 (既存処理に任せる)
            # confirm_suggestion がバインドされているためここでは何もしないか、検索実行
            pass
        else:
            # リスト操作
            self._activate_cursor_item(open_pdf=pdf_mode)
            return "break"

    def _handle_space_key(self, event):
        """Spaceキーのハンドラ"""
        if self._is_search_entry_focused():
            # 検索バーなら文字入力 (スペース) なので何もしない
            pass
        else:
            # リスト選択切り替え
            self._toggle_cursor_selection()
            return "break"

    def _handle_home_key(self, event, range_select=False):
        """Enterキーのハンドラ"""
        if (
            self._is_search_entry_focused()
        ):  # 検索バーなら検索実行またはオートコンプリート確定
            return

        # リスト操作
        is_ctrl = self._is_ctrl_pressed(event)
        self._move_cursor_absolute(0, range_select, keep_existing=is_ctrl)
        return "break"

    def _handle_end_key(self, event, range_select=False):
        """Spaceキーのハンドラ"""
        if (
            self._is_search_entry_focused()
        ):  # 検索バーなら文字入力 (スペース) なので何もしない
            return

        is_ctrl = self._is_ctrl_pressed(event)
        self._move_cursor_absolute(-1, range_select, keep_existing=is_ctrl)
        return "break"

    def _handle_page_up_key(self, event, range_select=False):
        if self._is_search_entry_focused():
            return

        # リスト選択切り替え
        is_ctrl = self._is_ctrl_pressed(event)
        self._move_cursor_relative(-10, range_select, keep_existing=is_ctrl)
        return "break"

    def _handle_page_down_key(self, event, range_select=False):
        if self._is_search_entry_focused():
            return
        is_ctrl = self._is_ctrl_pressed(event)
        self._move_cursor_relative(10, range_select, keep_existing=is_ctrl)
        return "break"

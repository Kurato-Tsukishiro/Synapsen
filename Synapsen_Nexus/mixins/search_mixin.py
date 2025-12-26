import logging

logger = logging.getLogger("Nexus.search_Mixin")


class NexusSearchMixin:
    """
    検索バーの入力制御を担当するMixin
    """

    def setup_search_variables(self):
        """検索関連の変数を初期化"""
        self.search_timer = None

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

        # オートコンプリートのスケジュール処理を削除し、検索のみスケジュール
        self.schedule_search()

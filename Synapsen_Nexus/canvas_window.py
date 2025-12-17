import uuid
import configparser
import datetime
import json
import logging
import os
import re
import shutil
import sqlite3
import sys
import tempfile
import time
import tkinter as tk
from pathlib import Path
from tkinter import filedialog, messagebox
from typing import Dict

import customtkinter as ctk
import fitz  # PyMuPDF

# --- プロジェクト内モジュールのパス設定 ---
current_dir = Path(__file__).parent
root_dir = current_dir.parent
normalisierer_dir = root_dir / "Synapsen_Normalisierer"

if str(root_dir) not in sys.path:
    sys.path.append(str(root_dir))
if str(normalisierer_dir) not in sys.path:
    sys.path.append(str(normalisierer_dir))

# --- プロジェクト内モジュールのインポート ---
from logging_setup import setup_logging  # noqa: E402
from Synapsen_Nexus.utils import (  # noqa: E402
    _update_note_links,
    _extract_links,
    touch_note_timestamp,
)

try:
    from pdf_utils import (  # type: ignore
        add_metadata_to_clip,
        convert_document_to_pdf,
        embed_processing_flag,
        high_fidelity_flatten,
        normalize_pdf_to_papersize,
    )
    from theme import SemanticColors as Colors
except ImportError:
    print("Warning: pdf_utils import failed.")

# --- ロガー設定 ---
logger = logging.getLogger("Nexus")
if not logger.handlers:
    setup_logging("Synapsen_Nexus")


# ==============================================================================
# 定数定義
# ==============================================================================
DEFAULT_CANVAS_BG_DARK = Colors.BACKGROUND_DARK_HOLLOW
DEFAULT_CANVAS_BG_LIGHT = Colors.BACKGROUND_HOLLOW
DEFAULT_GRID_COLOR_DARK = Colors.blend_colors(
    "#FFFFFF", Colors.BACKGROUND_DARK_PANEL, 0.2
)
DEFAULT_GRID_COLOR_LIGHT = Colors.BACKGROUND_PANEL

SHAPE_COLORS = {
    "レッド": "#FF4500",
    "イエロー": "#FFD700",
    "ブルー": "#1E90FF",
    "グリーン": "#32CD32",
    "グレー": "#808080",
    "ブラック": "#000000",
}

STICKY_COLORS = [
    ("イエロー", "#f8e58c"),  # 淡黄
    ("ブルー", "#bbc8e6"),  # 淡藤色
    ("レッド", "#eebbcb"),  # 撫子色
    ("グリーン", "#c1d8ac"),  # 裏葉柳
    ("ホワイト", "#fbfaf5"),  # 生成り色
    ("グレー", "#adadad"),  # 薄鈍
]


# ==============================================================================
# 基底クラス
# ==============================================================================
class BaseSubWindow(ctk.CTkToplevel):
    """アイコン設定などの共通機能を持つToplevel基底クラス"""

    def __init__(self, parent, title="Window", geometry=None):
        super().__init__(parent)
        self.parent = parent
        self.title(title)
        self.configure(
            fg_color=(Colors.BACKGROUND_HOLLOW, Colors.BACKGROUND_DARK_HOLLOW)
        )
        if geometry:
            self.geometry(geometry)

        self._custom_icon_path = None
        # 親、もしくは親の親からアイコンパスを取得
        if hasattr(parent, "icon_path") and parent.icon_path:
            self._custom_icon_path = str(parent.icon_path)
        elif hasattr(parent, "_custom_icon_path") and parent._custom_icon_path:
            self._custom_icon_path = parent._custom_icon_path
        elif (
            hasattr(parent, "parent_app")
            and hasattr(parent.parent_app, "icon_path")
            and parent.parent_app.icon_path
        ):
            self._custom_icon_path = str(parent.parent_app.icon_path)

        if self._custom_icon_path:
            self.after(200, self._apply_icon)

    def _apply_icon(self):
        if self._custom_icon_path:
            try:
                self.iconbitmap(default=self._custom_icon_path)
            except Exception:
                pass

    def iconbitmap(self, *args, **kwargs):
        """アイコン設定のオーバーライド"""
        if self._custom_icon_path:
            try:
                super().iconbitmap(default=self._custom_icon_path)
            except Exception:
                pass
        else:
            try:
                super().iconbitmap(*args, **kwargs)
            except Exception:
                pass


# ==============================================================================
# ダイアログクラス
# ==============================================================================
class StickyNoteDialog(BaseSubWindow):
    def __init__(self, parent, title_val="", content_val="", color_val="#FFFFA5"):
        super().__init__(parent, title="付箋の編集", geometry="450x450")
        self.grab_set()
        self.result = None
        self.selected_color = tk.StringVar(value=color_val)

        self._create_widgets(title_val, content_val)

        # ショートカットキーの追加
        # Ctrl+Enter で確定 (保存)
        self.bind("<Control-Return>", lambda e: self.on_ok())
        # Esc でキャンセル (閉じる)
        self.bind("<Escape>", lambda e: self.destroy())

    def _create_widgets(self, title_val, content_val):
        self.grid_columnconfigure(0, weight=1)
        self.grid_rowconfigure(5, weight=1)

        # 色選択
        ctk.CTkLabel(self, text="色:", anchor="w").grid(
            row=0, column=0, padx=10, pady=(10, 2), sticky="ew"
        )
        color_frame = ctk.CTkFrame(self, fg_color="transparent")
        color_frame.grid(row=1, column=0, padx=10, pady=(0, 10), sticky="ew")

        text_color = "gray90" if ctk.get_appearance_mode() == "Dark" else "black"

        for i, (name, code) in enumerate(STICKY_COLORS):
            btn = ctk.CTkRadioButton(
                color_frame,
                text=name,
                value=code,
                variable=self.selected_color,
                fg_color=code,
                border_color=Colors.STATE_DISABLED,
                text_color=text_color,
                width=80,
            )
            r, c = divmod(i, 3)
            btn.grid(row=r, column=c, padx=5, pady=5, sticky="w")

        # テキスト入力
        ctk.CTkLabel(self, text="タイトル (改行不可):", anchor="w").grid(
            row=2, column=0, padx=10, pady=(5, 2), sticky="ew"
        )
        self.title_entry = ctk.CTkEntry(self)
        self.title_entry.grid(row=3, column=0, padx=10, pady=(0, 10), sticky="ew")
        self.title_entry.insert(0, title_val)

        ctk.CTkLabel(self, text="内容:", anchor="w").grid(
            row=4, column=0, padx=10, pady=(0, 2), sticky="ew"
        )
        self.content_textbox = ctk.CTkTextbox(
            self,
            fg_color=(
                Colors.adjust_brightness(Colors.BACKGROUND_HOLLOW, 1.2),
                Colors.adjust_brightness(Colors.BACKGROUND_DARK_HOLLOW, 1.2),
            ),
        )
        self.content_textbox.grid(row=5, column=0, padx=10, pady=(0, 10), sticky="nsew")
        self.content_textbox.insert("1.0", content_val)

        # ボタン
        btn_frame = ctk.CTkFrame(self, fg_color="transparent")
        btn_frame.grid(row=6, column=0, padx=10, pady=10, sticky="e")

        ctk.CTkButton(
            btn_frame,
            text="キャンセル",
            width=80,
            fg_color=Colors.UI_CANCEL,
            hover_color=Colors.adjust_brightness(Colors.UI_CANCEL),
            command=self.destroy,
        ).pack(side="left", padx=5)
        ctk.CTkButton(
            btn_frame,
            text="OK",
            width=80,
            fg_color=Colors.UI_BASIC,
            hover_color=Colors.adjust_brightness(Colors.UI_BASIC),
            text_color="black",
            command=self.on_ok,
        ).pack(side="left", padx=5)

    def on_ok(self):
        raw_title = self.title_entry.get().replace("\n", " ").replace("\r", "").strip()
        raw_content = self.content_textbox.get("1.0", "end-1c")
        color = self.selected_color.get()
        self.result = (raw_title, raw_content, color)
        self.destroy()


class ConversionDialog(BaseSubWindow):
    """タイトルとIndex Keyを指定するダイアログ"""

    def __init__(
        self,
        parent,
        default_title,
        key_options,
        initial_color="#FFFFA5",
        show_color_option=True,
    ):
        geometry = "500x350" if show_color_option else "500x250"
        super().__init__(parent, title="詳細設定", geometry=geometry)
        self.grab_set()
        self.result = None
        self.key_options = key_options
        self.show_color_option = show_color_option
        self.selected_color = tk.StringVar(value=initial_color)

        self._create_widgets(default_title)

    def _create_widgets(self, default_title):
        self.grid_columnconfigure(0, weight=1)
        current_row = 0

        # 1. タイトル入力
        ctk.CTkLabel(
            self, text="タイトル (日付時刻はファイル名に自動付与されます):", anchor="w"
        ).grid(row=current_row, column=0, padx=20, pady=(20, 5), sticky="ew")
        current_row += 1

        self.title_entry = ctk.CTkEntry(self)
        self.title_entry.grid(
            row=current_row, column=0, padx=20, pady=(0, 15), sticky="ew"
        )
        self.title_entry.insert(0, default_title)
        current_row += 1

        # 2. Index Key 選択
        ctk.CTkLabel(self, text="Index Key (分類):", anchor="w").grid(
            row=current_row, column=0, padx=20, pady=(0, 5), sticky="ew"
        )
        current_row += 1

        self.key_combo = ctk.CTkComboBox(self, values=self.key_options)
        self.key_combo.grid(
            row=current_row,
            column=0,
            padx=20,
            pady=(0, 15 if self.show_color_option else 20),
            sticky="ew",
        )
        if self.key_options:
            self.key_combo.set(self.key_options[0])
        current_row += 1

        # 3. 色選択 (オプション)
        if self.show_color_option:
            ctk.CTkLabel(self, text="背景色:", anchor="w").grid(
                row=current_row, column=0, padx=20, pady=(0, 5), sticky="ew"
            )
            current_row += 1

            color_frame = ctk.CTkFrame(self, fg_color="transparent")
            color_frame.grid(
                row=current_row, column=0, padx=20, pady=(0, 15), sticky="ew"
            )
            current_row += 1

            text_color = "gray90" if ctk.get_appearance_mode() == "Dark" else "black"

            for i, (name, code) in enumerate(STICKY_COLORS):
                btn = ctk.CTkRadioButton(
                    color_frame,
                    text=name,
                    value=code,
                    variable=self.selected_color,
                    fg_color=code,
                    border_color=Colors.STATE_DISABLED,
                    text_color=text_color,
                    width=80,
                )
                r, c = divmod(i, 3)
                btn.grid(row=r, column=c, padx=5, pady=5, sticky="w")

        # 4. ボタン
        btn_frame = ctk.CTkFrame(self, fg_color="transparent")
        btn_frame.grid(row=current_row, column=0, padx=20, pady=10, sticky="e")

        ctk.CTkButton(
            btn_frame,
            text="キャンセル",
            width=80,
            fg_color=Colors.UI_CANCEL,
            command=self.destroy,
        ).pack(side="left", padx=5)

        ok_text = "変換実行" if self.show_color_option else "PDF出力"
        ctk.CTkButton(btn_frame, text=ok_text, width=80, command=self.on_ok).pack(
            side="left", padx=5
        )

    def on_ok(self):
        title = self.title_entry.get().strip()
        if not title:
            title = "NOTITLE"
        key = self.key_combo.get()

        if self.show_color_option:
            color = self.selected_color.get()
            self.result = (title, key, color)
        else:
            self.result = (title, key)
        self.destroy()


class CanvasHelpWindow(BaseSubWindow):
    def __init__(self, parent):
        super().__init__(parent, title="Canvas ヘルプ", geometry="600x750")
        self.grab_set()
        self._create_content()

    def _create_content(self):
        help_text = """
■ キーボードショートカット
---------------------------
[ツール切り替え]
V : 選択/移動モード
C : 接続 (Connect) モード
T : 付箋 (Sticky) モード
R : 枠線 (Rect) モード
L : 線 (Line) モード

[アクション]
N : 選択中のノートを追加
Ctrl + A : すべて選択
Backspace / Delete : 選択したアイテムを削除
Ctrl + S : 保存
Ctrl + Z : 元に戻す (Undo)
Ctrl + Y : やり直し (Redo)
Ctrl + W : 閉じる (Nexusへ戻る)
Esc : 選択解除 / ツールリセット

[表示操作]
Ctrl + + : ズームイン
Ctrl + - : ズームアウト
Ctrl + 0 : ズームリセット (全体表示)

[移動操作]
矢印キー / WASD : アイテムを移動 (同時押しで斜め移動)
Shift + 移動 : 速度切り替え (減速/加速)
Shift (2回押し) : 速度モードの固定切り替え

[付箋ウィンドウ内]
Ctrl + Enter : 確定 (保存して閉じる)
Esc : キャンセル (保存せずに閉じる)

■ 基本操作
---------------------------
[選択ツール (V)]
・クリック: アイテムを選択
・ドラッグ: アイテムを移動
・リサイズ: 選択時、右下の「■ハンドル」をドラッグ
・Shift + クリック: 複数選択の追加/解除
・背景ドラッグ: 範囲選択

[キーボード移動]
・矢印キー または WASDキー で選択アイテムを移動できます。
・デフォルトは「高速移動 (10px)」です。
・Shiftキーを押している間は「精密移動 (1px)」になります。
・Shiftキーを素早く2回押すと、デフォルトの速度設定が反転（ロック）します。

[視点操作]
・マウスホイール: 縦スクロール
・Shift + ホイール: 横スクロール
・Ctrl + ホイール: ズームイン/アウト
・ツールバーの +/- ボタン: ズーム操作

■ ツール
---------------------------
[付箋 (T)]
・ドラッグまたはクリックで作成。
・ダブルクリックで再編集。
・右クリックメニューから「PDFを作成」が可能。
    付箋の内容を正規化済みPDFとして出力します。
    ※ 本文にMarkdown形式で記述すると、PDF出力時に見出しやリスト等が反映されます。
・サイズ変更可能。
    ※ PDF出力時、付箋のサイズからはみ出した文字は印字されません。
       長文を入力した場合は、必ずサイズを広げて文字が収まるように調整してください。


[接続 (C)]
・アイテム間をドラッグして接続。
・ダブルクリックでリンク作成/解除 (Nexus DBと連携)。

[枠 (R) / 線 (L)]
・装飾用の図形を描画。色はツールバーで変更可能。

■ 連携機能
---------------------------
[OR検索]: 選択したノートでNexus検索を実行
[ノート追加]: Nexusで選択中のノートをキャンバスに追加
[PDF出力]: キャンバス全体をPDFとして保存
"""
        textbox = ctk.CTkTextbox(
            self, wrap="word", font=("", 12), fg_color=Colors.BACKGROUND_PANEL
        )
        textbox.pack(fill="both", expand=True, padx=10, pady=10)
        textbox.insert("1.0", help_text)
        textbox.configure(state="disabled")


# ==============================================================================
# CanvasWindow (メインクラス)
# ==============================================================================
class CanvasWindow(BaseSubWindow):
    def __init__(self, parent_app, background=False, initial_notes=None):
        super().__init__(parent_app, title="Synapsen Canvas")

        self.background_mode = background
        self.initial_notes_queue = initial_notes  # 起動後に追加するノートのリスト

        self.parent_app = parent_app
        self.default_save_file = parent_app.base_path / "canvas_data.json"

        # --- 履歴管理用変数 ---
        self.undo_stack = []
        self.redo_stack = []
        self.MAX_HISTORY = 20  # メモリ重視で20回に設定

        # --- データ構造 ---
        self.notes_on_canvas: Dict[str, Dict] = {}
        self.stickies_on_canvas: list = []
        self.shapes_on_canvas: list = []
        self.connections_on_canvas: list = []

        self.selected_items = set()
        self.current_scale = 1.0
        self.base_font_size_note = 10
        self.base_font_size_sticky = 12
        self.base_line_width = 2

        # ズーム時のオフセット (原点移動)
        self.canvas_offset_x = 0.0
        self.canvas_offset_y = 0.0

        self.current_mode = "select"
        self.drag_data = {
            "x": 0,
            "y": 0,
            "start_x": 0,
            "start_y": 0,
            "temp_id": None,
            "start_item": None,
            "rubberband_id": None,
            "moved": False,
        }

        self.font_path = self._get_font_path_from_config()
        self.current_shape_color = "レッド"

        # --- 移動操作用の状態変数 ---
        self.pressed_keys = set()
        self._move_loop_running = False
        self.last_shift_release_time = 0.0
        self.is_speed_locked = True  # 加速固定モードのフラグ

        # --- UI構築 ---
        self._create_ui()

        # --- 初期化処理 ---
        if not self.background_mode:
            self.lift()
            # 最初だけ最前面にして気づかせ _perform_initial_setup()で解除
            self.attributes("-topmost", True)
        else:
            self.lower()

        self.after(100, self._perform_initial_setup)

        # ウィンドウサイズ確定待ちのため、少し遅延させてからロードを実行
        self.after(200, self._initial_load)

    def _perform_initial_setup(self):
        """
        ウィンドウ生成直後に一度だけ実行される初期設定。
        最大化やバックグラウンド状態の設定を行う。
        """
        try:
            # 1. まず最大化状態にする (全モード共通)
            self.state("zoomed")

            # 2. 最前面固定を解除
            self.attributes("-topmost", False)

            # 3. モードに応じたフォーカス制御
            if not self.background_mode:
                # 通常起動: フォーカスを当てる
                self.focus_force()
            else:
                # バックグラウンド起動: すぐに最小化
                self.state("iconic")
                if self.parent_app:
                    self.parent_app.focus_force()

        except Exception:
            pass

    def _initial_load(self):
        """起動時の初期データロード"""
        if self.default_save_file.exists():
            self.load_canvas_data(self.default_save_file)
        else:
            # データがない場合でもグリッドを描画してセンタリング
            self._draw_grid(width=5000, height=5000)
            self.center_view()

        # ロードと配置(center_view)が終わった後に、待機していたノートを追加
        if self.initial_notes_queue:
            self.add_selected_notes(target_keys=self.initial_notes_queue)
            self.initial_notes_queue = None  # キューを空にする

    def on_close(self):
        """ウィンドウを閉じる際の処理"""

        # 閉じる前に親ウィンドウ(Nexus)を最前面にしてフォーカスを返す
        if self.parent_app and self.parent_app.winfo_exists():
            # 最小化されていたら元に戻す
            if self.parent_app.state() == "iconic":
                self.parent_app.deiconify()

            self.parent_app.lift()
            self.parent_app.focus_force()

        self.destroy()

    def _create_ui(self):
        self.toolbar = ctk.CTkFrame(self, fg_color=Colors.BACKGROUND_PANEL)
        self.toolbar.pack(side="top", fill="x", padx=5, pady=5)

        # ファイル操作系
        self._add_tool_btn(
            "開く",
            50,
            self.load_from_file,
            fg=Colors.UI_SECONDARY,
            hover=Colors.adjust_brightness(Colors.UI_SECONDARY),
        )
        self._add_tool_btn("保存", 50, self.save_canvas)
        self._add_tool_btn("別名保存", 60, self.save_as_file)
        self._add_sep()

        self._add_tool_btn(
            "PDF出力",
            60,
            self.export_canvas_dialog,
            fg=Colors.UI_EXPORT,
            hover=Colors.adjust_brightness(Colors.UI_EXPORT),
        )
        self._add_tool_btn(
            "画像出力",
            60,
            self.export_canvas_image_dialog,
            fg=Colors.UI_EXPORT,
            hover=Colors.adjust_brightness(Colors.UI_EXPORT),
        )
        self._add_tool_btn(
            "Mermaid出力",
            80,
            self.export_canvas_mermaid_dialog,
            fg=Colors.UI_EXPORT,
            hover=Colors.adjust_brightness(Colors.UI_EXPORT),
        )

        # ズーム系
        self._add_sep(" ズーム:")
        self._add_tool_btn("-", 30, self._zoom_out_btn)
        self.zoom_label_var = ctk.StringVar(value="100%")
        self.zoom_reset_button = ctk.CTkButton(
            self.toolbar,
            textvariable=self.zoom_label_var,
            width=50,
            command=self.reset_view,
            font=("", 12),
            fg_color="transparent",
            border_width=1,
            border_color="gray",
            text_color=("black", "white"),
            hover_color=(
                Colors.adjust_brightness(Colors.BACKGROUND_PANEL),
                Colors.adjust_brightness(Colors.BACKGROUND_DARK_PANEL, 1.2),
            ),
        )
        self.zoom_reset_button.pack(side="left", padx=2)
        self._add_tool_btn("+", 30, self._zoom_in_btn)

        # 色・ツール
        self._add_sep(" 色:")
        self.color_var = ctk.StringVar(value="レッド")
        self.color_menu = ctk.CTkOptionMenu(
            self.toolbar,
            values=list(SHAPE_COLORS.keys()),
            variable=self.color_var,
            width=90,
            fg_color=Colors.UI_BASIC,
            button_color=Colors.adjust_brightness(Colors.UI_BASIC),
            button_hover_color=Colors.adjust_brightness(Colors.UI_BASIC, 0.6),
            text_color="black",
            dropdown_fg_color=Colors.BACKGROUND_PANEL,
            dropdown_hover_color=Colors.adjust_brightness(Colors.BACKGROUND_PANEL),
        )
        self.color_menu.pack(side="left", padx=2)

        self._add_sep(" ツール:")
        self.mode_buttons = {}
        self.status_label_var = ctk.StringVar(value="現在のツール: 選択/移動")
        self.status_label = ctk.CTkLabel(
            self.toolbar,
            textvariable=self.status_label_var,
            font=("", 12, "bold"),
            text_color="gray",
        )

        self._add_mode_btn("select", "選択", "選択/移動")
        self._add_mode_btn("connect", "⇔ 接続", "ノート接続 (リンク)")
        self._add_mode_btn("sticky", "付箋", "付箋作成")
        self._add_mode_btn("rect", "□ 枠", "枠線")
        self._add_mode_btn("line", "／ 線", "線 (Path)")

        self._add_sep()
        self._add_tool_btn(
            "OR検索",
            70,
            self.send_or_search_to_nexus,
            fg=Colors.CANVAS,
            hover=Colors.adjust_brightness(Colors.CANVAS),
            text_col="black",
        )

        # 右側ボタン
        ctk.CTkButton(
            self.toolbar,
            text="？",
            width=30,
            command=self.show_help,
            fg_color=Colors.UI_BASIC,
            hover_color=Colors.adjust_brightness(Colors.UI_BASIC),
            text_color="black",
        ).pack(side="right", padx=5)
        ctk.CTkButton(
            self.toolbar,
            text="ノート追加",
            command=self.add_selected_notes,
            fg_color=Colors.UI_BASIC,
            hover_color=Colors.adjust_brightness(Colors.UI_BASIC),
            text_color="black",
        ).pack(side="right", padx=5)
        ctk.CTkButton(
            self.toolbar,
            text="全消去",
            width=60,
            fg_color=Colors.LABEL_DENGER,
            hover_color=Colors.adjust_brightness(Colors.LABEL_DENGER),
            command=self.clear_canvas,
        ).pack(side="right", padx=5)

        self.status_label.pack(side="right", padx=15)
        self.set_mode("select", "選択/移動")

        # キャンバス
        self.canvas_frame = ctk.CTkFrame(self)
        self.canvas_frame.pack(fill="both", expand=True, padx=5, pady=5)

        is_dark = ctk.get_appearance_mode() == "Dark"
        self.bg_color = DEFAULT_CANVAS_BG_DARK if is_dark else DEFAULT_CANVAS_BG_LIGHT
        self.grid_color = (
            DEFAULT_GRID_COLOR_DARK if is_dark else DEFAULT_GRID_COLOR_LIGHT
        )

        self.canvas = tk.Canvas(
            self.canvas_frame, bg=self.bg_color, highlightthickness=0
        )
        self._setup_scrollbars()
        self.canvas.pack(fill="both", expand=True)

        self._draw_grid(width=5000, height=5000)
        self._bind_events()
        self._setup_shortcuts()
        self._setup_move_shortcuts()
        self._setup_undo_shortcuts()

    def _add_tool_btn(
        self,
        text,
        w,
        cmd,
        fg=Colors.UI_BASIC,
        hover=Colors.adjust_brightness(Colors.UI_BASIC),
        text_col=None,
    ):
        text_col = (
            "black" if ((fg is Colors.UI_BASIC) and (text_col is None)) else text_col
        )
        btn = ctk.CTkButton(
            self.toolbar,
            text=text,
            width=w,
            command=cmd,
            fg_color=fg,
            hover_color=hover,
        )
        if text_col:
            btn.configure(text_color=text_col)
        btn.pack(side="left", padx=2)

    def _add_sep(self, text="|"):
        ctk.CTkLabel(self.toolbar, text=text).pack(side="left", padx=5)

    def _add_mode_btn(self, mode, text, label):
        btn = ctk.CTkButton(
            self.toolbar,
            text=text,
            width=50,
            hover_color=Colors.adjust_brightness(Colors.STATE_NORMAL),
            command=lambda: self.set_mode(mode, label),
        )
        btn.pack(side="left", padx=2)
        self.mode_buttons[mode] = btn

    def _setup_scrollbars(self):
        hbar = tk.Scrollbar(
            self.canvas_frame, orient=tk.HORIZONTAL, command=self.canvas.xview
        )
        hbar.pack(side=tk.BOTTOM, fill=tk.X)
        vbar = tk.Scrollbar(
            self.canvas_frame, orient=tk.VERTICAL, command=self.canvas.yview
        )
        vbar.pack(side=tk.RIGHT, fill=tk.Y)
        self.canvas.config(xscrollcommand=hbar.set, yscrollcommand=vbar.set)

    def _bind_events(self):
        self.canvas.bind("<ButtonPress-1>", self.on_canvas_click)
        self.canvas.bind("<B1-Motion>", self.on_canvas_drag)
        self.canvas.bind("<ButtonRelease-1>", self.on_canvas_release)
        self.canvas.bind("<Double-Button-1>", self.on_double_click)
        self.canvas.bind("<Button-3>", self.on_right_click)
        self.canvas.bind("<MouseWheel>", self._on_mousewheel)
        self.canvas.bind("<Shift-MouseWheel>", self._on_shift_mousewheel)
        self.canvas.bind("<Control-MouseWheel>", self._on_zoom_wheel)

    def _setup_shortcuts(self):
        """キャンバス操作用のキーボードショートカットを設定"""

        # --- ツール切り替え ---
        # V: 選択 (Select / Move)
        self.bind("v", lambda e: self.set_mode("select", "選択/移動"))
        self.bind("V", lambda e: self.set_mode("select", "選択/移動"))

        # C: 接続 (Connect)
        self.bind("c", lambda e: self.set_mode("connect", "ノート接続 (リンク)"))
        self.bind("C", lambda e: self.set_mode("connect", "ノート接続 (リンク)"))

        # T: 付箋 (Tag / Text / Sticky)
        self.bind("t", lambda e: self.set_mode("sticky", "付箋作成"))
        self.bind("T", lambda e: self.set_mode("sticky", "付箋作成"))

        # R: 枠線 (Rectangle)
        self.bind("r", lambda e: self.set_mode("rect", "枠線"))
        self.bind("R", lambda e: self.set_mode("rect", "枠線"))

        # L: 線 (Line)
        self.bind("l", lambda e: self.set_mode("line", "線 (Path)"))
        self.bind("L", lambda e: self.set_mode("line", "線 (Path)"))

        # --- ズーム操作 ---
        # Ctrl + + : ズームイン
        # (JISキーボードの「;」キーやテンキー、US配列の「=」キーなど複数のパターンに対応)
        self.bind("<Control-plus>", lambda e: self.zoom(1.1))
        self.bind("<Control-semicolon>", lambda e: self.zoom(1.1))
        self.bind("<Control-equal>", lambda e: self.zoom(1.1))
        self.bind("<Control-KP_Add>", lambda e: self.zoom(1.1))

        # Ctrl + - : ズームアウト
        self.bind("<Control-minus>", lambda e: self.zoom(0.9))
        self.bind("<Control-KP_Subtract>", lambda e: self.zoom(0.9))

        # Ctrl + 0 : ズームリセット (フィット表示)
        self.bind("<Control-0>", lambda e: self.reset_view())

        # --- アクション ---
        # N: ノート追加 (Nexusで選択中のノートを追加)
        self.bind("n", lambda e: self.add_selected_notes())
        self.bind("N", lambda e: self.add_selected_notes())

        # Ctrl+S: 保存
        self.bind("<Control-s>", lambda e: self.save_canvas())
        self.bind("<Control-S>", lambda e: self.save_canvas())

        # Ctrl+W: 閉じる (Nexusに戻る)
        self.bind("<Control-w>", lambda e: self.on_close())
        self.bind("<Control-W>", lambda e: self.on_close())

        # Ctrl + A : すべて選択
        self.bind("<Control-a>", lambda e: self.select_all_items())
        self.bind("<Control-A>", lambda e: self.select_all_items())

        # Esc: キャンセル / 選択モードへ復帰
        self.bind("<Escape>", self._handle_escape)

        # Deleate: 選択削除
        self.bind("<Delete>", self.delete_selected_items)
        self.bind("<BackSpace>", self.delete_selected_items)  # エイリアス

    def _setup_move_shortcuts(self):
        """
        移動関連のショートカットを設定
        """

        move_keys = [
            "Up",
            "Down",
            "Left",
            "Right",
            "w",
            "a",
            "s",
            "d",
            "W",
            "A",
            "S",
            "D",
            "Shift_L",
            "Shift_R",
        ]

        for k in move_keys:
            # 押した瞬間と離した瞬間をフックする
            self.bind(f"<KeyPress-{k}>", self._on_key_press)
            self.bind(f"<KeyRelease-{k}>", self._on_key_release)

    def _setup_undo_shortcuts(self):
        """Undo/Redoのショートカットを設定"""
        self.bind("<Control-z>", self.undo)
        self.bind("<Control-Z>", self.undo)
        self.bind("<Control-y>", self.redo)
        self.bind("<Control-Y>", self.redo)
        # Shift+Z も Redo として扱う (一般的な挙動)
        self.bind("<Control-Shift-z>", self.redo)
        self.bind("<Control-Shift-Z>", self.redo)

    def select_all_items(self):
        """キャンバス上のすべてのアイテムを選択状態にする"""
        self._clear_selection()

        # ノートを選択
        for key in self.notes_on_canvas:
            self.selected_items.add(("note", key))

        # 付箋を選択
        for s in self.stickies_on_canvas:
            self.selected_items.add(("sticky", s["uid"]))

        # 図形を選択
        for s in self.shapes_on_canvas:
            self.selected_items.add(("shape", s["uid"]))

        self._update_selection_visuals()

    def _handle_escape(self, event):
        """
        Escキーの挙動:
        1. もし「選択モード以外」なら -> 選択モードに戻す
        2. もし「選択モード」で「アイテム選択中」なら -> 選択解除
        """
        if self.current_mode != "select":
            self.set_mode("select", "選択/移動")
        else:
            # 選択モードですでに選択アイテムがある場合は解除
            if self.selected_items:
                self._clear_selection()

    def set_mode(self, mode, label_text):
        self.current_mode = mode
        self.status_label_var.set(f"現在のツール: {label_text}")
        for m, btn in self.mode_buttons.items():
            btn.configure(
                fg_color=(
                    Colors.adjust_brightness(Colors.STATE_NORMAL, 0.6)
                    if m == mode
                    else Colors.adjust_brightness(Colors.STATE_DISABLED, 1.2)
                )
            )
        if mode != "select":
            self._clear_selection()

    def show_help(self):
        if hasattr(self, "help_window") and self.help_window.winfo_exists():
            self.help_window.focus()
            return
        self.help_window = CanvasHelpWindow(self)
        self.help_window.focus()

    def _get_font_path_from_config(self):
        try:
            config_path = self.parent_app.base_path / "config.ini"
            if not config_path.is_file():
                config_path = self.parent_app.base_path.parent / "config.ini"
            if config_path.is_file():
                config = configparser.ConfigParser()
                config.read(config_path, encoding="utf-8")
                font_path_str = config.get("Paths", "font_path", fallback="")
                return Path(os.path.expandvars(font_path_str))
        except Exception:
            pass
        return None

    def _get_config_value(self, section, key, fallback):
        try:
            config_path = self.parent_app.base_path / "config.ini"
            if not config_path.is_file():
                config_path = self.parent_app.base_path.parent / "config.ini"
            if config_path.is_file():
                config = configparser.ConfigParser()
                config.read(config_path, encoding="utf-8")
                return config.get(section, key, fallback=fallback)
        except Exception:
            pass
        return fallback

    # -------------------------------------------------------------------------
    # Undo / Redo 機能の実装
    # -------------------------------------------------------------------------

    def _snapshot(self):
        """
        操作直前の状態を履歴に保存する。
        """
        # 現在の状態を軽量な辞書データとして取得
        current_state = self._generate_save_data_dict()

        self.undo_stack.append(current_state)
        self.redo_stack.clear()  # 新しい操作をしたらRedoは消える

        # 履歴制限を超えたら古いものを削除 (メモリ保護)
        if len(self.undo_stack) > self.MAX_HISTORY:
            self.undo_stack.pop(0)

    def undo(self, event=None):
        """元に戻す"""
        if not self.undo_stack:
            return

        # 現在の状態をRedoスタックに退避
        current_state = self._generate_save_data_dict()
        self.redo_stack.append(current_state)

        # 過去の状態を復元
        prev_state = self.undo_stack.pop()
        self._restore_from_dict(prev_state)
        self.save_canvas()  # ファイルにも反映

    def redo(self, event=None):
        """やり直し"""
        if not self.redo_stack:
            return

        # 現在の状態をUndoスタックに戻す
        current_state = self._generate_save_data_dict()
        self.undo_stack.append(current_state)

        # 未来の状態を復元
        next_state = self.redo_stack.pop()
        self._restore_from_dict(next_state)
        self.save_canvas()

    def _restore_from_dict(self, data):
        """辞書データから画面を再描画する"""
        self.clear_canvas_items(reset_view=False)  # 画面クリア

        # 各アイテムの復元 (load_canvas_data のロジックを流用)
        for k, v in data.get("notes", {}).items():
            self.create_note_item(
                k,
                v["title"],
                v.get("cp_key", ""),
                v["x"],
                v["y"],
                w=v.get("w", 160),
                h=v.get("h", 80),
            )

        for s in data.get("stickies", []):
            self.create_sticky_item(
                s["x"],
                s["y"],
                s.get("title", ""),
                s.get("content", ""),
                bg_color=s.get("bg_color", "#FFFFA5"),
                w=s.get("w", 180),
                h=s.get("h", 120),
                uid=s.get("uid"),
            )

        for s in data.get("shapes", []):
            if s["type"] == "line":
                self.create_shape_item(
                    "line",
                    x=s["x1"],
                    y=s["y1"],
                    x2=s["x2"],
                    y2=s["y2"],
                    color=s.get("color", "black"),
                    uid=s.get("uid"),
                )
            elif s["type"] == "rect":
                self.create_shape_item(
                    "rect",
                    x=s["x"],
                    y=s["y"],
                    w=s.get("w", 0),
                    h=s.get("h", 0),
                    color=s.get("color", "red"),
                    uid=s.get("uid"),
                )

        for c in data.get("connections", []):
            ft = c.get("from_type", "note")
            tt = c.get("to_type", "note")
            self.create_connection_item((ft, c["from_key"]), (tt, c["to_key"]))

        self.canvas.configure(scrollregion=self.canvas.bbox("all"))
        self.update_idletasks()

    # -------------------------------------------------------------------------
    # 方向キーによる移動の実装
    # -------------------------------------------------------------------------

    def _on_key_press(self, event):
        """キーが押されたらセットに追加し、移動ループを開始"""
        keysym = event.keysym
        self.pressed_keys.add(keysym)

        if not self._move_loop_running:
            # 移動開始の瞬間に一度だけスナップショット（Undo履歴）を作成
            self._snapshot()
            self._move_loop_running = True
            self._movement_loop()

    def _on_key_release(self, event):
        """キーが離されたらセットから削除 + Shiftダブルタップ判定"""
        keysym = event.keysym

        if keysym in self.pressed_keys:
            self.pressed_keys.remove(keysym)

        # --- Shiftキーのダブルタップ判定 (追加) ---
        if keysym in ("Shift_L", "Shift_R"):
            current_time = time.time()
            # 0.3秒以内に再度離されたらダブルタップとみなす
            if current_time - self.last_shift_release_time < 0.3:
                self.is_speed_locked = not self.is_speed_locked

                # 視覚的なフィードバック (オプション: ステータスバー等に表示)
                status = "ON" if self.is_speed_locked else "OFF"
                self.status_label_var.set(f"高速移動ロック: {status}")

                # ダブルタップ成立後は時間をリセット (トリプルタップでの誤爆防止)
                self.last_shift_release_time = 0.0
            else:
                self.last_shift_release_time = current_time
        # ----------------------------------------

        # 全ての移動キーが離されたら、変更をファイルに保存
        if not self.pressed_keys:
            self.save_canvas()

    def _movement_loop(self):
        """
        一定間隔でキーの状態を確認し、移動を実行するループ
        """
        if not self.pressed_keys:
            self._move_loop_running = False
            return

        dx = 0
        dy = 0

        # Shiftキーが押されているか
        is_shift_pressed = (
            "Shift_L" in self.pressed_keys
            or "Shift_R" in self.pressed_keys
            or any(k.isupper() for k in self.pressed_keys if len(k) == 1)
        )

        # ロック状態とShift状態が「異なる」場合に加速する (XOR)
        #  - ロックON (True)  & Shiftなし (False) => 加速 (True)  [デフォルト高速]
        #  - ロックON (True)  & Shiftあり (True)  => 減速 (False) [Shiftで減速]
        #  - ロックOFF(False) & Shiftなし (False) => 減速 (False) [デフォルト低速]
        #  - ロックOFF(False) & Shiftあり (True)  => 加速 (True)  [Shiftで加速]
        is_accelerated = self.is_speed_locked != is_shift_pressed

        step = 10 if is_accelerated else 1

        # 上下左右の判定
        if (
            "Up" in self.pressed_keys
            or "w" in self.pressed_keys
            or "W" in self.pressed_keys
        ):
            dy -= step
        if (
            "Down" in self.pressed_keys
            or "s" in self.pressed_keys
            or "S" in self.pressed_keys
        ):
            dy += step
        if (
            "Left" in self.pressed_keys
            or "a" in self.pressed_keys
            or "A" in self.pressed_keys
        ):
            dx -= step
        if (
            "Right" in self.pressed_keys
            or "d" in self.pressed_keys
            or "D" in self.pressed_keys
        ):
            dx += step

        if dx != 0 or dy != 0:
            self._nudge_selection_no_snapshot(dx, dy)

        # 約30ms後に再実行
        self.after(30, self._movement_loop)

    def _nudge_selection_no_snapshot(self, dx, dy):
        """nudge_selectionから _snapshot() を除いた軽量版"""
        if not self.selected_items:
            return

        ldx = dx / self.current_scale
        ldy = dy / self.current_scale

        for type_, key in self.selected_items:
            # ノート
            if type_ == "note":
                info = self.notes_on_canvas[key]
                self.canvas.move(info["ids"][0], dx, dy)
                self.canvas.move(info["ids"][1], dx, dy)
                info["x"] += ldx
                info["y"] += ldy
                hid = self.canvas.find_withtag(f"handle_note_{key}")
                if hid:
                    self.canvas.move(hid, dx, dy)
                self.update_connections(type_, key)

            # 付箋
            elif type_ == "sticky":
                s = next((x for x in self.stickies_on_canvas if x["uid"] == key), None)
                if s:
                    self.canvas.move(s["ids"][0], dx, dy)
                    self.canvas.move(s["ids"][1], dx, dy)
                    self.canvas.move(s["shadow_id"], dx, dy)
                    s["x"] += ldx
                    s["y"] += ldy
                    hid = self.canvas.find_withtag(f"handle_sticky_{key}")
                    if hid:
                        self.canvas.move(hid, dx, dy)
                    self.update_connections(type_, key)

            # 図形
            elif type_ == "shape":
                s = next((x for x in self.shapes_on_canvas if x["uid"] == key), None)
                if s:
                    self.canvas.move(s["id"], dx, dy)
                    if s["type"] == "rect":
                        s["x"] += ldx
                        s["y"] += ldy
                        hid = self.canvas.find_withtag(f"handle_shape_{key}")
                        if hid:
                            self.canvas.move(hid, dx, dy)
                    elif s["type"] == "line":
                        s["x1"] += ldx
                        s["y1"] += ldy
                        s["x2"] += ldx
                        s["y2"] += ldy

    # --- 描画関連 ---
    def _draw_grid(self, width, height, step=50):
        self.canvas.delete("grid")
        scaled_step = step * self.current_scale

        # 通常グリッド
        for x in range(0, int(width * self.current_scale), int(scaled_step)):
            self.canvas.create_line(
                x,
                0,
                x,
                height * self.current_scale,
                fill=self.grid_color,
                tags=("grid",),
            )
        for y in range(0, int(height * self.current_scale), int(scaled_step)):
            self.canvas.create_line(
                0,
                y,
                width * self.current_scale,
                y,
                fill=self.grid_color,
                tags=("grid",),
            )

        # 用紙サイズガイド
        self._draw_page_guide(width, height)

        self.canvas.tag_lower("grid")
        self.canvas.configure(scrollregion=self.canvas.bbox("all"))

    def _draw_page_guide(self, width, height):
        base_pw = getattr(self.parent_app, "paper_width", 595.276)
        base_ph = getattr(self.parent_app, "paper_height", 841.89)
        p_size_name = self.parent_app.config_data.get("paper_size", "A4")

        # x4サイズ
        pw, ph = base_pw * 4, base_ph * 4
        cx, cy = width / 2, height / 2

        x1 = (cx - pw / 2) * self.current_scale
        y1 = (cy - ph / 2) * self.current_scale
        x2 = (cx + pw / 2) * self.current_scale
        y2 = (cy + ph / 2) * self.current_scale

        guide_color = (
            Colors.LABEL_DENGER
            if ctk.get_appearance_mode() == "Dark"
            else Colors.LABEL_DENGER
        )

        self.canvas.create_rectangle(
            x1,
            y1,
            x2,
            y2,
            outline=guide_color,
            width=2,
            dash=(10, 10),
            tags=("grid", "guide"),
        )
        self.canvas.create_text(
            x1 + 10,
            y1 + 10,
            text=f"{p_size_name} (x4) ガイド",
            anchor="nw",
            fill=guide_color,
            font=("", int(12 * self.current_scale)),
            tags=("grid", "guide"),
        )

        # 十字
        cross_size = 20 * self.current_scale
        ccx, ccy = cx * self.current_scale, cy * self.current_scale
        self.canvas.create_line(
            ccx - cross_size,
            ccy,
            ccx + cross_size,
            ccy,
            fill=guide_color,
            width=2,
            tags=("grid", "guide"),
        )
        self.canvas.create_line(
            ccx,
            ccy - cross_size,
            ccx,
            ccy + cross_size,
            fill=guide_color,
            width=2,
            tags=("grid", "guide"),
        )

    # --- ズーム/スクロール操作 ---
    def _on_mousewheel(self, event):
        self.canvas.yview_scroll(int(-1 * (event.delta / 120)), "units")

    def _on_shift_mousewheel(self, event):
        self.canvas.xview_scroll(int(-1 * (event.delta / 120)), "units")

    def _zoom_in_btn(self):
        self.zoom(1.1)

    def _zoom_out_btn(self):
        self.zoom(0.9)

    def _on_zoom_wheel(self, event):
        if event.delta > 0:
            self.zoom(1.1, center_x=event.x, center_y=event.y)
        else:
            self.zoom(0.9, center_x=event.x, center_y=event.y)

    def center_view(self):
        self.update_idletasks()
        try:
            sr = self.canvas.cget("scrollregion").split()
            if not sr:
                return
            x1, y1, x2, y2 = map(float, sr)
            content_w, content_h = x2 - x1, y2 - y1
            if content_w <= 0 or content_h <= 0:
                return

            view_w = self.canvas.winfo_width()
            view_h = self.canvas.winfo_height()

            x_fraction = ((content_w / 2) - (view_w / 2)) / content_w
            y_fraction = ((content_h / 2) - (view_h / 2)) / content_h

            self.canvas.xview_moveto(x_fraction)
            self.canvas.yview_moveto(y_fraction)
        except Exception as e:
            logger.error(f"Centering failed: {e}")

    def reset_view(self):
        if self.current_scale == 0:
            return
        target_scale = 1.0
        if 0.99 < self.current_scale < 1.01:
            target_scale = 0.3
        scale_factor = target_scale / self.current_scale
        self.zoom(scale_factor)
        self.center_view()

    def zoom(self, scale_factor, center_x=None, center_y=None):
        if center_x is None:
            center_x = self.canvas.winfo_width() / 2
        if center_y is None:
            center_y = self.canvas.winfo_height() / 2
        canvas_x = self.canvas.canvasx(center_x)
        canvas_y = self.canvas.canvasy(center_y)
        self.canvas.scale("all", canvas_x, canvas_y, scale_factor, scale_factor)
        self.current_scale *= scale_factor

        self.canvas_offset_x = self.canvas_offset_x * scale_factor + canvas_x * (
            1 - scale_factor
        )
        self.canvas_offset_y = self.canvas_offset_y * scale_factor + canvas_y * (
            1 - scale_factor
        )

        self.zoom_label_var.set(f"{int(self.current_scale * 100)}%")
        self._apply_zoom_style()
        self.canvas.configure(scrollregion=self.canvas.bbox("all"))

    def _apply_zoom_style(self):
        new_note_font = ("", max(6, int(self.base_font_size_note * self.current_scale)))
        new_sticky_font = (
            "",
            max(8, int(self.base_font_size_sticky * self.current_scale)),
        )
        base_width = max(1, int(self.base_line_width * self.current_scale))
        selected_width = max(2, int(4 * self.current_scale))

        # ノート更新
        for key, info in self.notes_on_canvas.items():
            rect_id, text_id = info["ids"]
            self.canvas.itemconfigure(text_id, font=new_note_font)
            is_sel = ("note", key) in self.selected_items
            self.canvas.itemconfigure(
                rect_id,
                width=selected_width if is_sel else base_width,
                outline=Colors.UI_SECONDARY if is_sel else "white",
            )
            current_w = info["w"] * self.current_scale
            self.canvas.itemconfigure(
                text_id, width=current_w - (10 * self.current_scale)
            )

        # 付箋更新
        for sticky in self.stickies_on_canvas:
            rect_id, text_id = sticky["ids"]
            self.canvas.itemconfigure(text_id, font=new_sticky_font)

            is_sel = ("sticky", sticky["uid"]) in self.selected_items

            self.canvas.itemconfigure(
                rect_id,
                width=selected_width if is_sel else 0,
                outline=Colors.UI_SECONDARY if is_sel else "",
            )
            current_w = sticky["w"] * self.current_scale
            self.canvas.itemconfigure(
                text_id, width=current_w - (20 * self.current_scale)
            )

        # シェイプ更新
        for s in self.shapes_on_canvas:
            if s["type"] in ("rect", "line"):
                is_sel = ("shape", s["uid"]) in self.selected_items

                w = selected_width if is_sel else base_width
                base_color = s.get(
                    "color",
                    (
                        "red"
                        if s["type"] == "rect"
                        else (
                            "white" if ctk.get_appearance_mode() == "Dark" else "black"
                        )
                    ),
                )
                if s["type"] == "rect":
                    c = Colors.UI_SECONDARY if is_sel else base_color
                    self.canvas.itemconfigure(s["id"], outline=c, width=w)
                else:
                    c = Colors.UI_SECONDARY if is_sel else base_color
                    self.canvas.itemconfigure(s["id"], fill=c, width=w)

        for c in self.connections_on_canvas:
            self.canvas.itemconfigure(c["id"], width=base_width)

    # --- Selection Helpers ---
    def _get_item_key(self, item_type, data):
        if isinstance(data, str):
            return data
        if item_type == "note":
            return data  # ノートは key がそのままID
        if item_type in ("sticky", "shape"):
            return data.get("uid")
        return None

    def _select_item(self, item_type, item_data, add=False):
        if not add:
            self._clear_selection()
        key = self._get_item_key(item_type, item_data)
        if key:
            self.selected_items.add((item_type, key))
        self._update_selection_visuals()

    def _toggle_selection(self, item_type, item_data):
        key = self._get_item_key(item_type, item_data)
        if not key:
            return
        val = (item_type, key)
        if val in self.selected_items:
            self.selected_items.remove(val)
        else:
            self.selected_items.add(val)
        self._update_selection_visuals()

    def _clear_selection(self):
        self.selected_items.clear()
        self._update_selection_visuals()

    def _update_selection_visuals(self):
        self.canvas.delete("resize_handle")
        base_width = max(1, int(self.base_line_width * self.current_scale))
        selected_width = max(2, int(4 * self.current_scale))
        sel_col = Colors.UI_SECONDARY
        handle_size = 10 * self.current_scale

        # ノート (Note)
        for key, info in self.notes_on_canvas.items():
            rect_id = info["ids"][0]
            if ("note", key) in self.selected_items:
                self.canvas.itemconfigure(
                    rect_id, outline=sel_col, width=selected_width
                )
                self.canvas.tag_raise(rect_id)
                self.canvas.tag_raise(info["ids"][1])
                self._draw_handle(rect_id, handle_size, f"handle_note_{key}")
            else:
                self.canvas.itemconfigure(rect_id, outline="white", width=base_width)

        # 付箋 (Sticky)
        for sticky in self.stickies_on_canvas:
            rect_id = sticky["ids"][0]
            s_key = sticky["uid"]

            if ("sticky", s_key) in self.selected_items:
                self.canvas.itemconfigure(
                    rect_id, outline=sel_col, width=selected_width
                )
                self.canvas.tag_raise(rect_id)
                self.canvas.tag_raise(sticky["ids"][1])
                self._draw_handle(rect_id, handle_size, f"handle_sticky_{s_key}")
            else:
                self.canvas.itemconfigure(rect_id, outline="", width=0)

        # 図形 (Shape)
        for s in self.shapes_on_canvas:
            shape_key = s["uid"]
            is_sel = ("shape", shape_key) in self.selected_items

            if s["type"] == "rect":
                w = selected_width if is_sel else base_width
                base_color = s.get("color", "red")
                c = sel_col if is_sel else base_color
                self.canvas.itemconfigure(s["id"], outline=c, width=w)
                if is_sel:
                    self._draw_handle(s["id"], handle_size, f"handle_shape_{shape_key}")

            elif s["type"] == "line":
                w = selected_width if is_sel else base_width
                base_color = s.get(
                    "color", "white" if ctk.get_appearance_mode() == "Dark" else "black"
                )
                c = sel_col if is_sel else base_color
                self.canvas.itemconfigure(s["id"], fill=c, width=w)

    def _draw_handle(self, item_id, size, tag):
        coords = self.canvas.coords(item_id)
        if len(coords) == 4:
            x2, y2 = coords[2], coords[3]
            self.canvas.create_rectangle(
                x2 - size,
                y2 - size,
                x2,
                y2,
                fill=Colors.UI_SETTING,
                outline="white",
                tags=("resize_handle", tag),
            )

    def send_or_search_to_nexus(self):
        keys = [k for t, k in self.selected_items if t == "note"]
        if not keys:
            messagebox.showinfo(
                "情報", "検索対象のノートが選択されていません。", parent=self
            )
            return
        query_string = " OR ".join([f"[[{key}]]" for key in keys])
        self.parent_app.search_entry.delete(0, "end")
        self.parent_app.search_entry.insert(0, query_string)
        self.parent_app.perform_search()
        self.parent_app.lift()
        self.parent_app.focus_force()

    # --- アイテム作成 ---
    def add_selected_notes(self, target_keys=None):
        """Nexusで選択中のノートをキャンバスに追加する"""
        # 引数がなければ Nexus の選択状態を参照
        selected_keys = (
            target_keys if target_keys is not None else self.parent_app.selected_keys
        )

        if not selected_keys:
            # メッセージ表示などは状況によるが、Canvas内操作('N')なら表示しても良い
            if target_keys is None:
                messagebox.showinfo(
                    "情報", "Nexus本体でノートを選択してください。", parent=self
                )
            return

        self._snapshot()

        conn = self.parent_app.db_conn
        if not conn:
            return

        try:
            # 選択されたキーのリスト化
            keys_list = list(selected_keys)

            # SQLクエリの構築
            placeholders = ",".join("?" * len(keys_list))
            sql = (
                "SELECT key, title, "
                f"commonplace_key FROM notes WHERE key IN ({placeholders})"
            )

            cursor = conn.cursor()
            cursor.execute(sql, keys_list)
            rows = cursor.fetchall()

            # 配置座標の計算
            start_x = (
                self.canvas.canvasx(100) - self.canvas_offset_x
            ) / self.current_scale
            start_y = (
                self.canvas.canvasy(100) - self.canvas_offset_y
            ) / self.current_scale
            added = 0

            for row in rows:
                # DBから取得したタプル (key, title, commonplace_key)
                # sqlite3.Row factoryが設定されている場合もありますが、インデックスアクセスは共通で使えます
                key = row[0]
                title = row[1]
                cp_key = row[2]

                # 既にキャンバスにある場合はスキップ
                if key in self.notes_on_canvas:
                    continue

                # ノートを作成
                self.create_note_item(
                    key, title, cp_key if cp_key else "", start_x, start_y
                )

                # 座標をずらす
                start_x += 30
                start_y += 30
                added += 1

            if added > 0:
                self.save_canvas()

        except Exception as e:
            logger.error(f"Canvas add notes error: {e}")
            messagebox.showerror(
                "エラー", f"ノートの追加に失敗しました: {e}", parent=self
            )

    def create_note_item(self, key, title, cp_key, x, y, w=160, h=80):
        color = self.parent_app.key_colors.get(cp_key.lower(), "#aaaaaa")
        sw, sh = w * self.current_scale, h * self.current_scale
        sx = x * self.current_scale + self.canvas_offset_x
        sy = y * self.current_scale + self.canvas_offset_y
        fs = max(6, int(self.base_font_size_note * self.current_scale))
        lw = max(1, int(self.base_line_width * self.current_scale))

        rid = self.canvas.create_rectangle(
            sx,
            sy,
            sx + sw,
            sy + sh,
            fill=color,
            outline="white",
            width=lw,
            tags=("note", key),
        )

        dt = f"[[{key}: {title}]]"

        tid = self.canvas.create_text(
            sx + sw / 2,
            sy + sh / 2,
            text=dt,
            width=sw - (10 * self.current_scale),
            fill="black",
            font=("", fs),
            tags=("note", key),
            justify="center",
        )
        self.notes_on_canvas[key] = {
            "x": x,
            "y": y,
            "w": w,
            "h": h,
            "title": title,
            "cp_key": cp_key,
            "ids": (rid, tid),
        }

    def _create_sticky_display_text(self, title, content):
        return f"{title}\n{content}"

    def create_sticky_item(
        self, x, y, title="", content="", bg_color="#FFFFA5", w=180, h=120, uid=None
    ):
        if not uid:
            uid = uuid.uuid4().hex

        sw, sh = w * self.current_scale, h * self.current_scale
        sx = x * self.current_scale + self.canvas_offset_x
        sy = y * self.current_scale + self.canvas_offset_y
        pad = 10 * self.current_scale
        fs = max(8, int(self.base_font_size_sticky * self.current_scale))
        display_text = self._create_sticky_display_text(title, content)

        sid = self.canvas.create_rectangle(
            sx + 5,
            sy + 5,
            sx + sw + 5,
            sy + sh + 5,
            fill="gray50",
            outline="",
            tags=("sticky_shadow",),
        )
        rid = self.canvas.create_rectangle(
            sx,
            sy,
            sx + sw,
            sy + sh,
            fill=bg_color,
            outline="",
            width=0,
            tags=("sticky",),
        )
        tid = self.canvas.create_text(
            sx + pad,
            sy + pad,
            text=display_text,
            width=sw - (pad * 2),
            anchor="nw",
            font=("", fs),
            fill="black",
            tags=("sticky",),
        )

        item = {
            "uid": uid,
            "x": x,
            "y": y,
            "w": w,
            "h": h,
            "title": title,
            "content": content,
            "bg_color": bg_color,
            "ids": (rid, tid),
            "shadow_id": sid,
        }

        real_tag = f"sticky_{uid}"
        self.canvas.addtag_withtag(real_tag, rid)
        self.canvas.addtag_withtag(real_tag, tid)
        self.canvas.addtag_withtag(real_tag, sid)
        self.stickies_on_canvas.append(item)
        return item

    def create_connection_item(self, from_item, to_item):
        f_type, f_key = from_item
        t_type, t_key = to_item
        # 重複チェック
        for c in self.connections_on_canvas:
            if (
                c["from_key"] == f_key
                and c["to_key"] == t_key
                and c["from_type"] == f_type
                and c["to_type"] == t_type
            ) or (
                c["from_key"] == t_key
                and c["to_key"] == f_key
                and c["from_type"] == t_type
                and c["to_type"] == f_type
            ):
                return

        coords = self._get_connection_coords(f_type, f_key, t_type, t_key)
        if not coords:
            return
        x1, y1, x2, y2 = coords

        is_linked = False
        if f_type == "note" and t_type == "note":
            is_linked = self._check_db_link_exists(f_key, t_key)

        col = (
            Colors.UI_LINK
            if is_linked
            else ("white" if ctk.get_appearance_mode() == "Dark" else "black")
        )
        lw = max(1, int(self.base_line_width * self.current_scale))

        lid = self.canvas.create_line(
            x1, y1, x2, y2, fill=col, width=lw, arrow=tk.LAST, tags=("connection",)
        )
        self._lower_connection(lid)

        self.connections_on_canvas.append(
            {
                "id": lid,
                "from_key": f_key,
                "to_key": t_key,
                "from_type": f_type,
                "to_type": t_type,
            }
        )

    def _lower_connection(self, lid):
        for tag in ["note", "sticky"]:
            try:
                self.canvas.tag_lower(lid, tag)
            except tk.TclError:
                pass
        try:
            self.canvas.tag_raise(lid, "grid")
        except tk.TclError:
            self.canvas.tag_lower(lid)

    def _get_connection_coords(self, f_type, f_key, t_type, t_key):
        def get_bbox(type_, key_):
            if type_ == "note":
                if key_ in self.notes_on_canvas:
                    return self.canvas.bbox(self.notes_on_canvas[key_]["ids"][0])
            elif type_ == "sticky":
                for s in self.stickies_on_canvas:
                    if s["uid"] == key_:
                        return self.canvas.bbox(s["ids"][0])
            return None

        b1 = get_bbox(f_type, f_key)
        b2 = get_bbox(t_type, t_key)
        if not b1 or not b2:
            return None
        return (
            (b1[0] + b1[2]) / 2,
            (b1[1] + b1[3]) / 2,
            (b2[0] + b2[2]) / 2,
            (b2[1] + b2[3]) / 2,
        )

    def update_connections(self, type_, key_):
        for c in self.connections_on_canvas:
            if (c["from_type"] == type_ and c["from_key"] == key_) or (
                c["to_type"] == type_ and c["to_key"] == key_
            ):
                coords = self._get_connection_coords(
                    c["from_type"], c["from_key"], c["to_type"], c["to_key"]
                )
                if coords:
                    self.canvas.coords(c["id"], *coords)

    def create_shape_item(
        self, shape_type, x=0, y=0, w=0, h=0, text="", x2=0, y2=0, color="red", uid=None
    ):
        if not uid:
            uid = uuid.uuid4().hex

        lw = max(1, int(self.base_line_width * self.current_scale))
        sx = x * self.current_scale + self.canvas_offset_x
        sy = y * self.current_scale + self.canvas_offset_y
        sw, sh = w * self.current_scale, h * self.current_scale
        sx2 = x2 * self.current_scale + self.canvas_offset_x
        sy2 = y2 * self.current_scale + self.canvas_offset_y

        if shape_type == "rect":
            iid = self.canvas.create_rectangle(
                sx,
                sy,
                sx + sw,
                sy + sh,
                outline=color,
                width=lw,
                dash=(4, 4),
                tags=("shape", "rect"),
            )
            self.shapes_on_canvas.append(
                {
                    "uid": uid,
                    "id": iid,
                    "type": "rect",
                    "x": x,
                    "y": y,
                    "w": w,
                    "h": h,
                    "color": color,
                }
            )
            self.canvas.addtag_withtag(f"shape_{uid}", iid)

        elif shape_type == "line":
            iid = self.canvas.create_line(
                sx, sy, sx2, sy2, fill=color, width=lw, tags=("shape", "line")
            )
            self.shapes_on_canvas.append(
                {
                    "uid": uid,
                    "id": iid,
                    "type": "line",
                    "x1": x,
                    "y1": y,
                    "x2": x2,
                    "y2": y2,
                    "color": color,
                }
            )
            self.canvas.addtag_withtag(f"shape_{uid}", iid)

    def _check_db_link_exists(self, k1, k2):
        if not self.parent_app.loaded_db_path:
            return False
        conn = None
        try:
            conn = sqlite3.connect(
                f"file:{self.parent_app.loaded_db_path}?mode=ro", uri=True
            )
            cur = conn.cursor()
            cur.execute(
                "SELECT 1 FROM note_links WHERE "
                "(source_key=? AND target_key=?) "
                "OR (source_key=? AND target_key=?) LIMIT 1",
                (k1, k2, k2, k1),
            )
            return cur.fetchone() is not None
        except Exception:
            return False
        finally:
            if conn:
                conn.close()

    # --- Event Handlers (Refactored) ---
    def on_canvas_click(self, event):
        self.drag_data["snapshot_taken"] = False

        cx, cy = self.canvas.canvasx(event.x), self.canvas.canvasy(event.y)
        self.drag_data.update(
            {"x": cx, "y": cy, "start_x": cx, "start_y": cy, "moved": False}
        )

        # リサイズハンドル判定
        if self._handle_resize_click(cx, cy):
            return

        # アイテム判定
        target = self._find_target_item(cx, cy)
        is_shift = (event.state & 0x0001) != 0

        if self.current_mode == "connect":
            self._handle_connect_click(target, cx, cy)
        elif self.current_mode == "select":
            self._handle_select_click(target, cx, cy, is_shift)
        elif self.current_mode == "sticky":
            self._handle_create_sticky(cx, cy)
        elif self.current_mode in ("rect", "line"):
            self._handle_create_shape_start(cx, cy)

    def _handle_resize_click(self, cx, cy):
        clicked = self.canvas.find_overlapping(cx - 1, cy - 1, cx + 1, cy + 1)
        for item in clicked:
            tags = self.canvas.gettags(item)
            if "resize_handle" in tags:
                for t in tags:
                    if t.startswith("handle_sticky_"):
                        self.drag_data["start_item"] = (
                            "sticky_resize",
                            t.replace("handle_sticky_", ""),
                        )
                        self.current_mode = "resize_sticky"
                        return True
                    elif t.startswith("handle_note_"):
                        self.drag_data["start_item"] = (
                            "note_resize",
                            t.replace("handle_note_", ""),
                        )
                        self.current_mode = "resize_note"
                        return True
                    elif t.startswith("handle_shape_"):
                        self.drag_data["start_item"] = (
                            "shape_resize",
                            t.replace("handle_shape_", ""),
                        )
                        self.current_mode = "resize_shape"
                        return True
        return False

    def _find_target_item(self, cx, cy):
        # クリック判定
        items = self.canvas.find_overlapping(cx - 2, cy - 2, cx + 2, cy + 2)

        # 重なり順の逆（手前）から判定
        for item in reversed(items):
            tags = self.canvas.gettags(item)

            # ノート、付箋、図形の判定
            if "note" in tags:
                for t in tags:
                    if t not in ("note", "current"):
                        return ("note", t)
            elif "sticky" in tags:
                for s in self.stickies_on_canvas:
                    if s["ids"][0] == item or s["ids"][1] == item:
                        return ("sticky", s["uid"])
            elif "shape" in tags:
                for s in self.shapes_on_canvas:
                    if s["id"] == item:
                        return ("shape", s["uid"])
            elif "connection" in tags:
                return ("connection", item)

        return None

    def _handle_connect_click(self, target, cx, cy):
        if target and target[0] in ("note", "sticky"):
            self.drag_data["start_item"] = target
            col = "white" if ctk.get_appearance_mode() == "Dark" else "black"
            self.drag_data["temp_id"] = self.canvas.create_line(
                cx, cy, cx, cy, fill=col, width=2, dash=(2, 2)
            )

    def _handle_select_click(self, target, cx, cy, is_shift):
        if target:
            if is_shift:
                self._toggle_selection(*target)
            else:
                if target not in self.selected_items:
                    self._select_item(*target)
            self.drag_data["start_item"] = target
        else:
            if not is_shift:
                self._clear_selection()
            self.drag_data["rubberband_id"] = self.canvas.create_rectangle(
                cx, cy, cx, cy, outline=Colors.UI_SECONDARY, dash=(2, 2)
            )

    def _handle_create_sticky(self, cx, cy):
        dialog = StickyNoteDialog(self)
        self.wait_window(dialog)
        if dialog.result:
            title, content, color = dialog.result
            if title or content:
                self._snapshot()
                lx = (cx - self.canvas_offset_x) / self.current_scale
                ly = (cy - self.canvas_offset_y) / self.current_scale
                self.create_sticky_item(lx, ly, title, content, bg_color=color)
                self.save_canvas()

    def _handle_create_shape_start(self, cx, cy):
        current_color = SHAPE_COLORS.get(self.color_var.get(), "red")
        if self.current_mode == "rect":
            self.drag_data["temp_id"] = self.canvas.create_rectangle(
                cx, cy, cx, cy, outline=current_color
            )
        elif self.current_mode == "line":
            self.drag_data["temp_id"] = self.canvas.create_line(
                cx, cy, cx, cy, fill=current_color
            )

    def on_canvas_drag(self, event):
        # まだスナップショットをとっていない、かつ、操作モードの場合
        if not self.drag_data.get("snapshot_taken", False):
            # リサイズ中、または アイテム選択中の移動なら保存
            if self.current_mode in (
                "resize_sticky",
                "resize_note",
                "resize_shape",
            ) or (self.current_mode == "select" and self.selected_items):
                self._snapshot()  # ★ここで履歴保存
                self.drag_data["snapshot_taken"] = True

        cx, cy = self.canvas.canvasx(event.x), self.canvas.canvasy(event.y)
        dx, dy = cx - self.drag_data["x"], cy - self.drag_data["y"]
        self.drag_data.update({"moved": True, "x": cx, "y": cy})

        if self.current_mode in ("resize_sticky", "resize_note", "resize_shape"):
            self._handle_drag_resize(dx, dy)
        elif self.current_mode == "connect":
            self._handle_drag_connect(cx, cy)
        elif self.current_mode == "select":
            self._handle_drag_move(dx, dy, cx, cy)
        elif self.current_mode in ("rect", "line"):
            self._handle_drag_create_shape(cx, cy)

    def _handle_drag_resize(self, dx, dy):
        target_type, target_key = self.drag_data["start_item"]
        min_size = 30 * self.current_scale
        handle_size = 10 * self.current_scale

        target = None
        rect_id = None
        text_id = None  # For note/sticky

        if self.current_mode == "resize_sticky":
            target = next(
                (s for s in self.stickies_on_canvas if s["uid"] == target_key), None
            )
            if target:
                rect_id, text_id = target["ids"]
        elif self.current_mode == "resize_note":
            target = self.notes_on_canvas.get(target_key)
            if target:
                rect_id, text_id = target["ids"]
        elif self.current_mode == "resize_shape":
            target = next(
                (s for s in self.shapes_on_canvas if s["uid"] == target_key), None
            )
            if target and target["type"] == "rect":
                rect_id = target["id"]

        if target and rect_id:
            coords = self.canvas.coords(rect_id)
            x1, y1 = coords[0], coords[1]
            curr_w, curr_h = coords[2] - x1, coords[3] - y1
            new_w = max(min_size, curr_w + dx)
            new_h = max(min_size, curr_h + dy)

            target["w"] = new_w / self.current_scale
            target["h"] = new_h / self.current_scale

            self.canvas.coords(rect_id, x1, y1, x1 + new_w, y1 + new_h)

            if self.current_mode == "resize_sticky":
                self.canvas.coords(
                    target["shadow_id"],
                    x1 + 5,
                    y1 + 5,
                    x1 + new_w + 5,
                    y1 + new_h + 5,
                )
                pad = 10 * self.current_scale
                self.canvas.coords(text_id, x1 + pad, y1 + pad)
                self.canvas.itemconfigure(text_id, width=new_w - (pad * 2))
                hid = self.canvas.find_withtag(f"handle_sticky_{target_key}")

            elif self.current_mode == "resize_note":
                self.canvas.coords(text_id, x1 + new_w / 2, y1 + new_h / 2)
                self.canvas.itemconfigure(
                    text_id, width=new_w - (10 * self.current_scale)
                )
                hid = self.canvas.find_withtag(f"handle_note_{target_key}")

            elif self.current_mode == "resize_shape":
                hid = self.canvas.find_withtag(f"handle_shape_{target_key}")

            if hid:
                self.canvas.coords(
                    hid,
                    x1 + new_w - handle_size,
                    y1 + new_h - handle_size,
                    x1 + new_w,
                    y1 + new_h,
                )

    def _handle_drag_connect(self, cx, cy):
        if self.drag_data["temp_id"]:
            self.canvas.coords(
                self.drag_data["temp_id"],
                self.drag_data["start_x"],
                self.drag_data["start_y"],
                cx,
                cy,
            )

    def _handle_drag_move(self, dx, dy, cx, cy):
        ldx = dx / self.current_scale
        ldy = dy / self.current_scale

        if self.selected_items:
            for t, k in self.selected_items:
                if t == "note":
                    info = self.notes_on_canvas[k]
                    self.canvas.move(info["ids"][0], dx, dy)
                    self.canvas.move(info["ids"][1], dx, dy)
                    info["x"] += ldx
                    info["y"] += ldy
                    self.update_connections(t, k)
                    hid = self.canvas.find_withtag(f"handle_note_{k}")
                    if hid:
                        self.canvas.move(hid, dx, dy)

                elif t == "sticky":
                    s = next(
                        (x for x in self.stickies_on_canvas if x["uid"] == k), None
                    )
                    if s:
                        self.canvas.move(s["ids"][0], dx, dy)
                        self.canvas.move(s["ids"][1], dx, dy)
                        self.canvas.move(s["shadow_id"], dx, dy)
                        s["x"] += ldx
                        s["y"] += ldy
                        self.update_connections(t, k)
                        hid = self.canvas.find_withtag(f"handle_sticky_{k}")
                        if hid:
                            self.canvas.move(hid, dx, dy)

                elif t == "shape":
                    s = next((x for x in self.shapes_on_canvas if x["uid"] == k), None)
                    if s:
                        self.canvas.move(s["id"], dx, dy)
                        if s["type"] == "rect":
                            hid = self.canvas.find_withtag(f"handle_shape_{k}")
                            if hid:
                                self.canvas.move(hid, dx, dy)
                        if s["type"] == "line":
                            s["x1"] += ldx
                            s["y1"] += ldy
                            s["x2"] += ldx
                            s["y2"] += ldy
                        else:
                            s["x"] += ldx
                            s["y"] += ldy

        elif self.drag_data["rubberband_id"]:
            self.canvas.coords(
                self.drag_data["rubberband_id"],
                self.drag_data["start_x"],
                self.drag_data["start_y"],
                cx,
                cy,
            )

    def _handle_drag_create_shape(self, cx, cy):
        if self.drag_data["temp_id"]:
            current_color = SHAPE_COLORS.get(self.color_var.get(), "red")
            self.canvas.coords(
                self.drag_data["temp_id"],
                self.drag_data["start_x"],
                self.drag_data["start_y"],
                cx,
                cy,
            )
            if self.current_mode == "rect":
                self.canvas.itemconfigure(
                    self.drag_data["temp_id"], outline=current_color
                )
            else:
                self.canvas.itemconfigure(self.drag_data["temp_id"], fill=current_color)

    def on_canvas_release(self, event):
        cx, cy = self.canvas.canvasx(event.x), self.canvas.canvasy(event.y)

        if self.current_mode in ("resize_sticky", "resize_note", "resize_shape"):
            self.current_mode = "select"
            self.drag_data["start_item"] = None
            self.save_canvas()
            self._update_selection_visuals()
            return

        if self.current_mode == "select":
            if self.drag_data["rubberband_id"]:
                self._handle_rubberband_selection()
            elif self.drag_data["moved"]:
                self.save_canvas()

        elif self.current_mode == "connect":
            self._handle_connect_release(cx, cy)

        elif self.current_mode in ("rect", "line"):
            self._handle_create_shape_release(cx, cy)

        self.drag_data["start_item"] = None

    def _handle_rubberband_selection(self):
        bbox = self.canvas.coords(self.drag_data["rubberband_id"])
        if len(bbox) == 4:
            x1, y1, x2, y2 = (
                min(bbox[0], bbox[2]),
                min(bbox[1], bbox[3]),
                max(bbox[0], bbox[2]),
                max(bbox[1], bbox[3]),
            )
            enclosed = self.canvas.find_enclosed(x1, y1, x2, y2)
            for item in enclosed:
                tags = self.canvas.gettags(item)
                if "note" in tags:
                    for t in tags:
                        if t not in ("note", "current"):
                            self.selected_items.add(("note", t))
                elif "sticky" in tags:
                    for t in tags:
                        if t.startswith("sticky_"):
                            self.selected_items.add(("sticky", t.split("_")[1]))
                elif "shape" in tags:
                    for t in tags:
                        if t.startswith("shape_"):
                            self.selected_items.add(("shape", t.split("_")[1]))

        self.canvas.delete(self.drag_data["rubberband_id"])
        self.drag_data["rubberband_id"] = None
        self._update_selection_visuals()

    def _handle_connect_release(self, cx, cy):
        if self.drag_data["temp_id"]:
            self.canvas.delete(self.drag_data["temp_id"])
            self.drag_data["temp_id"] = None

        if self.drag_data["start_item"]:
            target = self._find_target_item(cx, cy)
            if target and target != self.drag_data["start_item"]:
                self.create_connection_item(self.drag_data["start_item"], target)
                self._snapshot()
                self.save_canvas()

    def _handle_create_shape_release(self, cx, cy):
        if self.drag_data["temp_id"]:
            self.canvas.delete(self.drag_data["temp_id"])
            if (
                abs(cx - self.drag_data["start_x"]) > 5
                or abs(cy - self.drag_data["start_y"]) > 5
            ):
                lx1 = (
                    self.drag_data["start_x"] - self.canvas_offset_x
                ) / self.current_scale
                ly1 = (
                    self.drag_data["start_y"] - self.canvas_offset_y
                ) / self.current_scale
                lx2 = (cx - self.canvas_offset_x) / self.current_scale
                ly2 = (cy - self.canvas_offset_y) / self.current_scale

                color = SHAPE_COLORS.get(self.color_var.get(), "red")

                if self.current_mode == "rect":
                    real_x, real_y = min(lx1, lx2), min(ly1, ly2)
                    w, h = abs(lx2 - lx1), abs(ly2 - ly1)
                    self.create_shape_item(
                        "rect", x=real_x, y=real_y, w=w, h=h, color=color
                    )
                    self._snapshot()
                elif self.current_mode == "line":
                    self.create_shape_item(
                        "line", x=lx1, y=ly1, x2=lx2, y2=ly2, color=color
                    )
                self.save_canvas()

    def on_double_click(self, event):
        cx, cy = self.canvas.canvasx(event.x), self.canvas.canvasy(event.y)
        item = self.canvas.find_closest(cx, cy)[0]
        tags = self.canvas.gettags(item)

        if "sticky" in tags:
            self._edit_sticky(item)
        elif "note" in tags:
            for t in tags:
                if t not in ("note", "current"):
                    self.parent_app.open_preview_window(t, ui_master=self)
                    break
        elif "connection" in tags:
            self._edit_connection(item)

    def _edit_sticky(self, item):
        target = None
        for s in self.stickies_on_canvas:
            if s["ids"][0] == item or s["ids"][1] == item:
                target = s
                break
        if target:
            dialog = StickyNoteDialog(
                self,
                title_val=target.get("title", ""),
                content_val=target.get("content", ""),
                color_val=target.get("bg_color", "#FFFFA5"),
            )
            self.wait_window(dialog)
            if dialog.result:
                self._snapshot()
                t, c, col = dialog.result
                target.update({"title": t, "content": c, "bg_color": col})
                display_text = self._create_sticky_display_text(t, c)
                self.canvas.itemconfigure(target["ids"][1], text=display_text)
                self.canvas.itemconfigure(target["ids"][0], fill=col)
                self.save_canvas()

    def _edit_connection(self, item):
        target_conn = next(
            (c for c in self.connections_on_canvas if c["id"] == item), None
        )
        if (
            target_conn
            and target_conn["from_type"] == "note"
            and target_conn["to_type"] == "note"
        ):
            self._handle_connection_double_click(target_conn)

    def on_right_click(self, event):
        cx, cy = self.canvas.canvasx(event.x), self.canvas.canvasy(event.y)
        target = self._find_target_item(cx, cy)

        if target:
            self._show_context_menu(event, target)

    def _show_context_menu(self, event, target):
        menu = tk.Menu(self, tearoff=0)
        type_, obj_id = target
        obj = None

        if type_ == "sticky":
            obj = next((x for x in self.stickies_on_canvas if x["uid"] == obj_id), None)
            if obj:
                menu.add_command(
                    label="PDFを作成 (保存のみ)",
                    command=lambda: self._convert_sticky_to_note_pipeline(obj),
                )
                selected_stickies = [
                    i[1] for i in self.selected_items if i[0] == "sticky"
                ]
                if len(selected_stickies) > 1 and obj["uid"] in selected_stickies:
                    menu.add_command(
                        label="まとめてPDFを作成",
                        command=self._convert_selected_stickies_pipeline,
                    )

        elif type_ == "connection":
            # 接続線の特定 (IDから検索)
            item = self.canvas.find_closest(
                self.canvas.canvasx(event.x), self.canvas.canvasy(event.y)
            )[0]
            # connections_on_canvasから該当する辞書を探す
            obj = next((c for c in self.connections_on_canvas if c["id"] == item), None)

            if obj and obj["from_type"] == "note" and obj["to_type"] == "note":
                is_linked = self._check_db_link_exists(obj["from_key"], obj["to_key"])
                label = "リンク解除" if is_linked else "リンク作成"
                menu.add_command(
                    label=label,
                    command=lambda: self._handle_connection_double_click(obj),
                )

        elif type_ == "shape":
            obj = next((x for x in self.shapes_on_canvas if x["uid"] == obj_id), None)
        elif type_ == "note":
            obj = obj_id  # note key

        # 削除メニューの表示ロジック
        if obj is not None or type_ == "note":
            # ★変更: 右クリックした対象が「選択中」かつ「複数選択されている」場合は一括削除を表示
            target_is_selected = (type_, obj_id) in self.selected_items

            if target_is_selected and len(self.selected_items) > 1:
                menu.add_separator()
                menu.add_command(
                    label=f"選択した {len(self.selected_items)} 項目を削除",
                    command=self.delete_selected_items,
                )
            else:
                menu.add_separator()
                menu.add_command(
                    label="削除", command=lambda: self._delete_item(type_, obj)
                )

            menu.post(event.x_root, event.y_root)

    # --- Connection Logic ---
    def _handle_connection_double_click(self, conn_data):
        key_a, key_b = conn_data["from_key"], conn_data["to_key"]
        title_a = self.notes_on_canvas[key_a]["title"]
        title_b = self.notes_on_canvas[key_b]["title"]

        if self._check_db_link_exists(key_a, key_b):
            if messagebox.askyesno(
                "リンク解除",
                f"リンク解除しますか？\n・{title_a}\n・{title_b}",
                parent=self,
            ):
                self._perform_db_link_removal(key_a, key_b, conn_data["id"])
        else:
            if messagebox.askyesno(
                "リンク作成",
                f"相互リンク作成しますか？\n・{title_a}\n・{title_b}",
                parent=self,
            ):
                self._perform_db_link_update(key_a, key_b, conn_data["id"])
        self.focus_force()

    def _perform_db_link_update(self, key_a, key_b, item_id):
        if not self.parent_app.loaded_db_path:
            return
        conn = None
        self._snapshot()
        try:
            conn = sqlite3.connect(self.parent_app.loaded_db_path)
            cursor = conn.cursor()

            # リンクテキストの追記
            self._append_link_text(
                cursor, key_a, key_b, self.notes_on_canvas[key_b]["title"]
            )
            self._append_link_text(
                cursor, key_b, key_a, self.notes_on_canvas[key_a]["title"]
            )

            touch_note_timestamp(cursor, key_a)
            touch_note_timestamp(cursor, key_b)

            conn.commit()
            self.canvas.itemconfigure(item_id, fill=Colors.UI_LINK)
            messagebox.showinfo("完了", "リンクを作成しました。", parent=self)
        except Exception as e:
            if conn:
                conn.rollback()
            messagebox.showerror("DBエラー", f"失敗: {e}", parent=self)
        finally:
            if conn:
                conn.close()

    def _perform_db_link_removal(self, key_a, key_b, item_id):
        if not self.parent_app.loaded_db_path:
            return
        conn = None
        self._snapshot()
        try:
            conn = sqlite3.connect(self.parent_app.loaded_db_path)
            cursor = conn.cursor()

            # リンクテキストの削除
            self._remove_link_text_from_db(cursor, key_a, key_b)
            self._remove_link_text_from_db(cursor, key_b, key_a)

            # 両方のノートの更新日時を更新
            touch_note_timestamp(cursor, key_a)
            touch_note_timestamp(cursor, key_b)

            conn.commit()
            base_color = "white" if ctk.get_appearance_mode() == "Dark" else "black"
            self.canvas.itemconfigure(item_id, fill=base_color)
            messagebox.showinfo("完了", "リンクを解除しました。", parent=self)
        except Exception as e:
            if conn:
                conn.rollback()
            messagebox.showerror("DBエラー", f"失敗: {e}", parent=self)
        finally:
            if conn:
                conn.close()

    def _append_link_text(self, cursor, target_key, link_key, link_title):
        cursor.execute("SELECT memo FROM notes WHERE key = ?", (target_key,))
        row = cursor.fetchone()
        current_memo = row[0] if row and row[0] else ""
        link_str = f"[[{link_key}: {link_title}]]"
        if link_str not in current_memo:
            new_memo = current_memo.strip() + f"\n{link_str}\n"
            cursor.execute(
                "UPDATE notes SET memo = ? WHERE key = ?", (new_memo, target_key)
            )
            _update_note_links(cursor, target_key, new_memo)

    def _remove_link_text_from_db(self, cursor, target_key, remove_link_key):
        cursor.execute("SELECT memo FROM notes WHERE key = ?", (target_key,))
        row = cursor.fetchone()
        current_memo = row[0] if row and row[0] else ""
        if not current_memo:
            return
        pattern = r"\[\[" + re.escape(remove_link_key) + r"(:.*?)?\]\]\s*"
        new_memo = re.sub(pattern, "", current_memo)
        if new_memo != current_memo:
            cursor.execute(
                "UPDATE notes SET memo = ? WHERE key = ?",
                (new_memo.strip(), target_key),
            )
            _update_note_links(cursor, target_key, new_memo.strip())

    # --- Export Helpers ---
    def _convert_sticky_to_note_pipeline(self, sticky_obj):
        default_title = sticky_obj.get("title", "NOTITLE")
        current_color = sticky_obj.get("bg_color", "#FFFFA5")
        key_options = getattr(self.parent_app, "commonplace_keys_options", [])

        dialog = ConversionDialog(
            self,
            default_title,
            key_options,
            initial_color=current_color,
            show_color_option=True,
        )
        self.wait_window(dialog)

        if dialog.result:
            title, key, color = dialog.result
            content = sticky_obj.get("content", "")

            # --- 引用Keyの自動収集 ---
            # 1. キャンバス上で接続されているノートのKey
            sticky_id = sticky_obj["uid"]
            connected_keys = self._get_connected_keys_for_item("sticky", sticky_id)

            # 2. 本文に含まれる [[Key]] リンク
            text_links = _extract_links(content)

            # 重複を除いてリスト化
            cited_keys = sorted(list(connected_keys | text_links))
            # -----------------------

            # 引数に cited_keys を追加して呼び出し
            self._process_md_pdf_creation(title, content, key, color, cited_keys)

    def _convert_selected_stickies_pipeline(self):
        combined_content = ""
        targets = []
        for t, k in self.selected_items:
            if t == "sticky":
                s = next((x for x in self.stickies_on_canvas if x["uid"] == k), None)
                if s:
                    targets.append(s)

        if not targets:
            return

        key_options = getattr(self.parent_app, "commonplace_keys_options", [])

        dialog = ConversionDialog(
            self,
            "付箋の結合",
            key_options,
            initial_color="#FFFFA5",
            show_color_option=True,
        )
        self.wait_window(dialog)

        if dialog.result:
            title, key, color = dialog.result

            all_cited_keys = set()  # 全ての付箋から引用Keyを収集するセット

            for s in targets:
                t_ = s.get("title", "NOTITLE")
                c_ = s.get("content", "")

                # コンテンツの結合
                combined_content += f"## {t_}\n{c_}\n\n"

                # 1. 本文内のリンクを抽出して追加
                all_cited_keys.update(_extract_links(c_))

                # 2. 各付箋に接続されているノートのKeyも収集して追加
                sticky_id = s["uid"]
                connected_keys = self._get_connected_keys_for_item("sticky", sticky_id)
                all_cited_keys.update(connected_keys)

            # 重複を除いてソート・リスト化
            cited_keys = sorted(list(all_cited_keys))

            # PDF生成処理へ渡す (cited_keysを追加)
            self._process_md_pdf_creation(
                title, combined_content, key, color, cited_keys
            )

    def _get_connected_keys_for_item(self, item_type, item_key):
        connected_keys = set()
        for c in self.connections_on_canvas:
            # この接続が対象アイテムを含んでいるかチェック
            other_type, other_key = None, None

            if c["from_type"] == item_type and c["from_key"] == item_key:
                other_type, other_key = c["to_type"], c["to_key"]
            elif c["to_type"] == item_type and c["to_key"] == item_key:
                other_type, other_key = c["from_type"], c["from_key"]

            # 接続相手が 'note' であれば、そのKeyを収集
            if other_type == "note" and other_key:
                connected_keys.add(other_key)

        return connected_keys

    def _process_md_pdf_creation(
        self, title, content, index_key, bg_color, cited_keys=None
    ):
        now_str = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
        safe_title = re.sub(r'[\\/:\*\?"<>\|]', "_", title if title else "NOTITLE")
        base_name = f"{now_str}_{safe_title}"

        save_dir = getattr(self.parent_app, "nexus_output_folder", Path("Nexus_Output"))
        if not save_dir.exists():
            try:
                save_dir.mkdir(parents=True, exist_ok=True)
            except Exception as e:
                messagebox.showerror(
                    "エラー", f"保存先を作成できません: {e}", parent=self
                )
                return

        pdf_path = save_dir / f"{base_name}.pdf"
        temp_dir = save_dir / "temp_canvas_process"
        temp_dir.mkdir(exist_ok=True)

        try:
            md_path = temp_dir / f"{base_name}.md"
            style_tag = (
                "<style>"
                "html {{ width: 100%; margin: 0; padding: 0; "
                f"background-color: {bg_color}; }}"
                f"body {{ background-color: {bg_color} !important; "
                "width: 100% !important; margin: 0 !important; "
                "padding: 0 !important; max-width: none !important; "
                "min-height: 100vh; }}"
                ".content-wrapper { padding: 20px; }"
                "</style>"
            )
            md_text = (
                f"{style_tag}\n\n"
                "<div class='content-wrapper'>\n\n"
                f"# {title}\n\n# [内容]\n{content}\n\n"
                "</div>"
            )
            with open(md_path, "w", encoding="utf-8") as f:
                f.write(md_text)

            temp_pdf = temp_dir / f"temp_{base_name}.pdf"
            temp_flat = temp_dir / f"flat_{base_name}.pdf"
            font_path = self.font_path or r"C:\Windows\Fonts\msgothic.ttc"

            convert_document_to_pdf(
                md_path,
                temp_pdf,
                paper_size_str="A4",
                pdf_margins={"top": "0", "bottom": "0", "left": "0", "right": "0"},
            )
            high_fidelity_flatten(
                str(temp_pdf), str(temp_flat), str(font_path), flatten_ink=False
            )
            normalize_pdf_to_papersize(
                str(temp_flat), str(pdf_path), 595.276, 841.89, target_format="A4"
            )
            try:
                doc = fitz.open(pdf_path)
                meta = doc.metadata
                # 既存のキーワードを取得し、識別子を追記
                current_keywords = meta.get("keywords", "")
                new_keywords = (
                    f"{current_keywords}; Synapsen:Sticky"
                    if current_keywords
                    else "Synapsen:Sticky"
                )
                meta["keywords"] = new_keywords
                doc.set_metadata(meta)
                doc.saveIncr()  # 増分保存
                doc.close()
            except Exception as e:
                logger.warning(f"付箋識別子の埋め込みに失敗: {e}")
            embed_processing_flag(str(pdf_path))

            # メタデータ埋め込み
            key_rect_tuple = (0, 13, 391, 73)  # Default
            text_color_rgb = (0, 0, 0)
            if index_key:
                hex_c = self.parent_app.key_colors.get(index_key.lower())
                if hex_c:
                    text_color_rgb = self._hex_to_rgb(hex_c)

            # configからQRサイズ取得
            refs_qr_size_str = self._get_config_value(
                "Extraction", "refs_qr_size", "75"
            )
            try:
                refs_qr_size_pt = int(refs_qr_size_str)
            except ValueError:
                refs_qr_size_pt = 75

            add_metadata_to_clip(
                pdf_path_str=str(pdf_path),
                font_path=str(font_path),
                paper_width=595.276,
                paper_height=841.89,
                key_rect_tuple=key_rect_tuple,
                index_key_to_embed=index_key,
                text_color=text_color_rgb,
                comment_to_embed=f"Sticky Note Export: {title}",
                base_name=base_name,
                cited_keys_list=cited_keys,
                refs_qr_size_pt=refs_qr_size_pt,
                extra_keywords=["Synapsen:Sticky"],
            )
            messagebox.showinfo(
                "完了", f"ファイルを生成しました:\n{pdf_path.name}", parent=self
            )
        except Exception as e:
            messagebox.showerror("エラー", f"処理に失敗しました:\n{e}", parent=self)
        finally:
            shutil.rmtree(temp_dir, ignore_errors=True)

    # -------------------------------------------------------------------------
    # Mermaid Export Implementation
    # -------------------------------------------------------------------------
    def export_canvas_mermaid_dialog(self):
        """Canvasの状態をMermaid記法としてクリップボードにコピー"""
        try:
            mermaid_text = self._generate_mermaid_text()

            # クリップボードへコピー
            self.clipboard_clear()
            self.clipboard_append(mermaid_text)
            self.update()

            messagebox.showinfo(
                "Mermaid出力",
                "Mermaid記法をクリップボードにコピーしました。\n"
                "NotionなどのMarkdown対応ツールに貼り付けて利用できます。\n\n"
                "(Notionの場合: /mermaid と入力してコードブロックを作成し、貼り付けてください)",
                parent=self,
            )
        except Exception as e:
            logger.error(f"Mermaid export error: {e}")
            messagebox.showerror("エラー", f"出力に失敗しました:\n{e}", parent=self)

    def _generate_mermaid_text(self):
        """現在のCanvasデータからMermaid構文を生成する"""
        lines = ["%% Synapsen Canvas Export", "graph TD"]

        # クラス定義 (色分け用)
        lines.append("    classDef note fill:#ffffff,stroke:#333,stroke-width:2px;")
        lines.append(
            "    classDef sticky_yellow fill:#FFFFA5,stroke:#e0e0e0,stroke-width:1px;"
        )
        lines.append(
            "    classDef sticky_blue fill:#D1EAFF,stroke:#e0e0e0,stroke-width:1px;"
        )
        lines.append(
            "    classDef sticky_red fill:#FFD1D1,stroke:#e0e0e0,stroke-width:1px;"
        )
        lines.append(
            "    classDef sticky_green fill:#D1FFD1,stroke:#e0e0e0,stroke-width:1px;"
        )
        lines.append(
            "    classDef sticky_gray fill:#E0E0E0,stroke:#e0e0e0,stroke-width:1px;"
        )

        # IDのサニタイズ用ヘルパー
        def clean_id(raw_id):
            # MermaidのIDに使用できない文字を除去
            return "ID_" + re.sub(r"[^a-zA-Z0-9_]", "", str(raw_id))

        def escape_text(text):
            # ラベル内の特殊文字をエスケープ
            return text.replace('"', "'").replace("\n", "<br>")

        # 1. ノート (Nodes)
        for key, info in self.notes_on_canvas.items():
            nid = clean_id(key)
            title = escape_text(info["title"])
            lines.append(f'    {nid}("{title}<br><small>{key}</small>"):::note')

        # 2. 付箋 (Stickies)
        color_map = {
            "#f8e58c": "sticky_yellow",
            "#bbc8e6": "sticky_blue",
            "#eebbcb": "sticky_red",
            "#c1d8ac": "sticky_green",
            "#fbfaf5": "sticky_gray",
            "#adadad": "note",
        }

        for s in self.stickies_on_canvas:
            sid = clean_id(s["uid"])
            title = escape_text(s.get("title", ""))
            content = escape_text(s.get("content", ""))

            # 長文の切り詰め
            if len(content) > 50:
                content = content[:50] + "..."

            display_text = f"<b>{title}</b>"
            if content:
                display_text += f"<br>{content}"

            bg_color = s.get("bg_color", "#f8e58c")
            style_class = color_map.get(bg_color, "sticky_yellow")

            lines.append(f'    {sid}["{display_text}"]:::{style_class}')

        # 3. 接続 (Connections)
        for c in self.connections_on_canvas:
            from_id = None
            to_id = None

            # ID解決
            if c["from_type"] == "note":
                from_id = clean_id(c["from_key"])
            elif c["from_type"] == "sticky":
                from_id = clean_id(c["from_key"])

            if c["to_type"] == "note":
                to_id = clean_id(c["to_key"])
            elif c["to_type"] == "sticky":
                to_id = clean_id(c["to_key"])

            if from_id and to_id:
                # DBリンク確認
                is_linked = False
                # 両端がノートの場合のみDBリンクの可能性があるためチェック
                if c["from_type"] == "note" and c["to_type"] == "note":
                    is_linked = self._check_db_link_exists(c["from_key"], c["to_key"])

                # 矢印のスタイル決定
                # 実線 (-->) : DBリンク済み（強い結合）
                # 点線 (-.->): DBリンクなし、または付箋との接続（弱い結合・一時的）
                arrow = "-->" if is_linked else "-.->"

                lines.append(f"    {from_id} {arrow} {to_id}")

        return "\n".join(lines)

    def _get_obj_from_key(self, type_, key):
        if type_ == "note":
            return key  # Noteの場合はKeyそのものがオブジェクト(ID)
        elif type_ == "sticky":
            return next((s for s in self.stickies_on_canvas if s["uid"] == key), None)
        elif type_ == "shape":
            return next((s for s in self.shapes_on_canvas if s["uid"] == key), None)
        return None

    def delete_selected_items(self, event=None):
        """選択中のアイテムを一括削除する"""
        self._snapshot()
        if not self.selected_items:
            return

        # 確認ダイアログ
        if not messagebox.askyesno(
            "削除確認",
            f"選択した {len(self.selected_items)} 個のアイテムを削除しますか？",
            parent=self,
        ):
            return

        # 削除ループ中にセットが変更されるのを防ぐためコピーを作成
        items_to_delete = list(self.selected_items)

        for type_, key in items_to_delete:
            obj = self._get_obj_from_key(type_, key)
            if obj is not None:
                # 個別の自動保存はスキップ (save_after=False)
                self._delete_item(type_, obj, save_after=False)

        # 選択状態をクリアして保存
        self._clear_selection()
        self.save_canvas()

    def _delete_item(self, type_, obj, save_after=True):
        if type_ == "sticky":
            self.canvas.delete(obj["ids"][0])
            self.canvas.delete(obj["ids"][1])
            self.canvas.delete(obj["shadow_id"])

            if obj in self.stickies_on_canvas:
                self.stickies_on_canvas.remove(obj)

            # 選択リストからも削除
            sid = obj["uid"]
            if ("sticky", sid) in self.selected_items:
                self.selected_items.remove(("sticky", sid))

            self._remove_associated_connections(sid, "sticky")

        elif type_ == "note":
            self._delete_note(obj)  # _delete_note内でselected_items操作あり
        elif type_ == "connection":
            self.canvas.delete(obj["id"])
            self.connections_on_canvas.remove(obj)
        elif type_ == "shape":
            self.canvas.delete(obj["id"])
            self.shapes_on_canvas.remove(obj)

            # 選択リストからも削除
            sid = obj["uid"]
            if ("shape", sid) in self.selected_items:
                self.selected_items.remove(("shape", sid))

        if save_after:
            self.save_canvas()

    def _delete_note(self, key):
        ids = self.notes_on_canvas[key]["ids"]
        for i in ids:
            self.canvas.delete(i)
        del self.notes_on_canvas[key]
        if ("note", key) in self.selected_items:
            self.selected_items.remove(("note", key))
        self._remove_associated_connections(key, "note")

    def _remove_associated_connections(self, key, type_):
        to_rem = [
            c
            for c in self.connections_on_canvas
            if (c["from_type"] == type_ and c["from_key"] == key)
            or (c["to_type"] == type_ and c["to_key"] == key)
        ]
        for c in to_rem:
            self.canvas.delete(c["id"])
            self.connections_on_canvas.remove(c)

    # --- File IO ---
    def save_canvas(self):
        self._save_to_json(self.default_save_file)

    def save_as_file(self):
        file_path = filedialog.asksaveasfilename(
            defaultextension=".json",
            filetypes=[("JSON Files", "*.json")],
            title="キャンバスを保存",
            parent=self,
        )
        if file_path:
            self._save_to_json(Path(file_path))
            messagebox.showinfo(
                "完了", f"保存しました: {Path(file_path).name}", parent=self
            )

    def _generate_save_data_dict(self):
        """
        現在のキャンバス状態を辞書データとして生成して返す。
        """
        data = {
            "notes": {
                k: {
                    "x": v["x"],
                    "y": v["y"],
                    "w": v["w"],
                    "h": v["h"],
                    "title": v["title"],
                    "cp_key": v["cp_key"],
                }
                for k, v in self.notes_on_canvas.items()
            },
            "stickies": [
                {
                    "uid": s["uid"],
                    "x": s["x"],
                    "y": s["y"],
                    "w": s["w"],
                    "h": s["h"],
                    "title": s.get("title", ""),
                    "content": s.get("content", ""),
                    "bg_color": s["bg_color"],
                }
                for s in self.stickies_on_canvas
            ],
            "shapes": [],
            "connections": [
                {
                    "from_key": c["from_key"],
                    "to_key": c["to_key"],
                    "from_type": c["from_type"],
                    "to_type": c["to_type"],
                }
                for c in self.connections_on_canvas
            ],
        }
        for s in self.shapes_on_canvas:
            item = s.copy()
            if "id" in item:
                del item["id"]  # IDは保存不要
            data["shapes"].append(item)
        return data

    def _save_to_json(self, path):
        # 共通化したデータ生成メソッドを使用
        data = self._generate_save_data_dict()
        try:
            with open(path, "w", encoding="utf-8") as f:
                json.dump(data, f, indent=2, ensure_ascii=False)
        except Exception as e:
            logger.error(f"Canvas save error: {e}")

    def load_from_file(self):
        file_path = filedialog.askopenfilename(
            filetypes=[("JSON Files", "*.json")], title="キャンバスを開く", parent=self
        )
        if file_path:
            if self.notes_on_canvas or self.stickies_on_canvas:
                if not messagebox.askyesno(
                    "確認", "現在のキャンバスをクリアして読み込みますか？", parent=self
                ):
                    return
            self.load_canvas_data(Path(file_path))

    def load_canvas_data(self, path):
        """キャンバスデータをファイルから読み込み、アイテムを再配置する"""
        if not path.exists():
            return

        # 読み込み前に一度完全にリセット (ズーム等も戻す)
        self.clear_canvas_items(reset_view=True)

        try:
            with open(path, "r", encoding="utf-8") as f:
                data = json.load(f)
            # 共通化した復元メソッドを使用
            self._restore_from_dict(data)
            self.undo_stack.clear()  # ロード時は履歴クリア
            self.redo_stack.clear()
            self.center_view()  # 最後に全体を表示

        except Exception as e:
            logger.error(f"Canvas load error: {e}")
            messagebox.showerror("エラー", f"読込失敗: {e}", parent=self)

    def clear_canvas(self):
        if messagebox.askyesno("確認", "キャンバスをクリアしますか？", parent=self):
            self.clear_canvas_items()
            self.center_view()
            self.save_canvas()

    def clear_canvas_items(self, reset_view=True):
        """
        キャンバス上のアイテムとデータを消去する。
        reset_view=True: 全消去して視点もリセット (新規作成、読込時)
        reset_view=False: グリッドと視点は残してコンテンツのみ消去 (Undo/Redo時)
        """
        if reset_view:
            # 完全に初期化する場合
            self.canvas.delete("all")

            # 視点のリセット
            self.current_scale = 1.0
            self.canvas_offset_x = 0.0
            self.canvas_offset_y = 0.0
            if hasattr(self, "zoom_label_var"):
                self.zoom_label_var.set("100%")

            self.canvas.xview_moveto(0)
            self.canvas.yview_moveto(0)

            # グリッド再描画
            self._draw_grid(width=5000, height=5000)

        else:
            # 視点を維持する場合 (Undo/Redo)
            # グリッド("grid", "guide")以外のタグを持つアイテムを削除する
            tags_to_delete = [
                "note",
                "sticky",
                "sticky_shadow",
                "shape",
                "connection",
                "resize_handle",
            ]
            for tag in tags_to_delete:
                self.canvas.delete(tag)

        # データ構造のリセット (共通)
        self.notes_on_canvas = {}
        self.stickies_on_canvas = []
        self.shapes_on_canvas = []
        self.connections_on_canvas = []
        self.selected_items = set()

    def export_canvas_dialog(self):
        """PDFとしてエクスポート"""
        key_options = getattr(self.parent_app, "commonplace_keys_options", [])
        dialog = ConversionDialog(
            self, "Canvas_Export", key_options, show_color_option=False
        )
        self.wait_window(dialog)

        if not dialog.result:
            return

        title_input, index_key_input = dialog.result
        now_str = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
        safe_title = re.sub(r'[\\/:\*\?"<>\|]', "_", title_input)
        initial_file = f"{now_str}_{safe_title}.pdf"

        out_dir = getattr(self.parent_app, "nexus_output_folder", None)
        initial_dir = str(out_dir) if out_dir and out_dir.exists() else None

        file_path = filedialog.asksaveasfilename(
            defaultextension=".pdf",
            filetypes=[("PDF Document", "*.pdf")],
            title="PDFとして保存",
            initialfile=initial_file,
            initialdir=initial_dir,
            parent=self,
        )
        if file_path:
            try:
                self._export_to_file(Path(file_path), title_input, index_key_input)
                messagebox.showinfo(
                    "完了", f"出力が完了しました:\n{Path(file_path).name}", parent=self
                )
            except Exception as e:
                messagebox.showerror("エラー", f"出力に失敗しました:\n{e}", parent=self)
                logger.error(f"PDF Export Error: {e}")

    def export_canvas_image_dialog(self):
        """画像(PNG/JPG)としてエクスポート"""
        # メタデータ(Title/IndexKey)は画像自体には埋め込めませんが、
        # ファイル名や将来的な拡張のために同じダイアログを使用します
        key_options = getattr(self.parent_app, "commonplace_keys_options", [])
        dialog = ConversionDialog(
            self, "Canvas_Image", key_options, show_color_option=False
        )
        self.wait_window(dialog)

        if not dialog.result:
            return

        title_input, index_key_input = dialog.result
        now_str = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
        safe_title = re.sub(r'[\\/:\*\?"<>\|]', "_", title_input)
        initial_file = f"{now_str}_{safe_title}.png"

        out_dir = getattr(self.parent_app, "nexus_output_folder", None)
        initial_dir = str(out_dir) if out_dir and out_dir.exists() else None

        file_path = filedialog.asksaveasfilename(
            defaultextension=".png",
            filetypes=[("PNG Image", "*.png"), ("JPEG Image", "*.jpg")],
            title="画像として保存",
            initialfile=initial_file,
            initialdir=initial_dir,
            parent=self,
        )
        if file_path:
            try:
                # 同じ _export_to_file を呼び出す (内部で拡張子を見て分岐)
                self._export_to_file(Path(file_path), title_input, index_key_input)
                messagebox.showinfo(
                    "完了", f"出力が完了しました:\n{Path(file_path).name}", parent=self
                )
            except Exception as e:
                messagebox.showerror("エラー", f"出力に失敗しました:\n{e}", parent=self)
                logger.error(f"Image Export Error: {e}")

    def _hex_to_rgb(self, hex_color):
        hex_color = hex_color.lstrip("#")
        return tuple(int(hex_color[i : i + 2], 16) / 255.0 for i in (0, 2, 4))

    def _export_to_file(self, output_path, title, index_key):
        # 1. コンテンツ領域の計算 (既存ロジック)
        min_x, min_y, max_x, max_y = (
            float("inf"),
            float("inf"),
            float("-inf"),
            float("-inf"),
        )

        def update_bounds(x, y, w, h):
            nonlocal min_x, min_y, max_x, max_y
            min_x, min_y = min(min_x, x), min(min_y, y)
            max_x, max_y = max(max_x, x + w), max(max_y, y + h)

        for info in self.notes_on_canvas.values():
            update_bounds(info["x"], info["y"], 160, 80)
        for s in self.stickies_on_canvas:
            update_bounds(s["x"], s["y"], s["w"], s["h"])
        for s in self.shapes_on_canvas:
            if s["type"] == "rect":
                update_bounds(s["x"], s["y"], s.get("w", 0), s.get("h", 0))
            elif s["type"] == "line":
                min_x = min(min_x, s["x1"], s["x2"])
                min_y = min(min_y, s["y1"], s["y2"])
                max_x = max(max_x, s["x1"], s["x2"])
                max_y = max(max_y, s["y1"], s["y2"])

        if min_x == float("inf"):
            min_x, min_y, max_x, max_y = 0, 0, 100, 100

        margin = 50
        content_w = max_x - min_x + margin * 2
        content_h = max_y - min_y + margin * 2

        # 2. PyMuPDFで描画 (PDF/画像 共通)
        doc = fitz.open()
        page = doc.new_page(width=content_w, height=content_h)

        # メタデータ設定 (PDF用)
        doc.set_metadata(
            {
                "keywords": "Synapsen:Whiteboard",
                "creator": "Synapsen Canvas",
                "title": title,
            }
        )

        font_path = self.font_path or r"C:\Windows\Fonts\msgothic.ttc"
        try:
            page.insert_font(fontname="embed_font", fontfile=str(font_path))
        except Exception:
            pass

        shape = page.new_shape()

        def tx(v):
            return v - min_x + margin

        def ty(v):
            return v - min_y + margin

        # ヘルパー: 中心座標取得
        def get_center(type_, key_):
            if type_ == "note":
                if key_ in self.notes_on_canvas:
                    info = self.notes_on_canvas[key_]
                    return tx(info["x"]) + 80, ty(info["y"]) + 40
            elif type_ == "sticky":
                s = next((x for x in self.stickies_on_canvas if x["uid"] == key_), None)
                if s:
                    return tx(s["x"]) + s["w"] / 2, ty(s["y"]) + s["h"] / 2
            return None, None

        # --- 描画実行 ---

        # 1. 接続線
        for c in self.connections_on_canvas:
            x1, y1 = get_center(c["from_type"], c["from_key"])
            x2, y2 = get_center(c["to_type"], c["to_key"])

            if x1 is not None and x2 is not None:
                is_linked = False
                if c["from_type"] == "note" and c["to_type"] == "note":
                    is_linked = self._check_db_link_exists(c["from_key"], c["to_key"])
                color = (0.15, 0.65, 0.27) if is_linked else (0, 0, 0)

                shape.draw_line(fitz.Point(x1, y1), fitz.Point(x2, y2))
                shape.finish(color=color, width=1)

        # 2. 図形
        for s in self.shapes_on_canvas:
            default_color = (
                "red"
                if s["type"] == "rect"
                else ("white" if ctk.get_appearance_mode() == "Dark" else "black")
            )
            color_val = s.get("color", default_color)
            color_map = {"red": "#FF0000", "white": "#FFFFFF", "black": "#000000"}

            if color_val in SHAPE_COLORS:
                hex_code = SHAPE_COLORS[color_val]
            else:
                hex_code = color_map.get(color_val, color_val)

            try:
                rgb_color = self._hex_to_rgb(hex_code)
            except Exception:
                rgb_color = (0, 0, 0)

            if s["type"] == "rect":
                rtx, rty = tx(s["x"]), ty(s["y"])
                shape.draw_rect(
                    fitz.Rect(rtx, rty, rtx + s.get("w", 0), rty + s.get("h", 0))
                )
                shape.finish(color=rgb_color, width=1, dashes=[4, 4])
            elif s["type"] == "line":
                x1, y1 = tx(s["x1"]), ty(s["y1"])
                x2, y2 = tx(s["x2"]), ty(s["y2"])
                shape.draw_line(fitz.Point(x1, y1), fitz.Point(x2, y2))
                shape.finish(color=rgb_color, width=1)

        # 3. 付箋
        for s in self.stickies_on_canvas:
            rtx, rty = tx(s["x"]), ty(s["y"])
            w, h = s["w"], s["h"]

            # 影
            shape.draw_rect(fitz.Rect(rtx + 5, rty + 5, rtx + w + 5, rty + h + 5))
            shape.finish(fill=(0.5, 0.5, 0.5), stroke_opacity=0)

            # 本体
            shape.draw_rect(fitz.Rect(rtx, rty, rtx + w, rty + h))
            shape.finish(fill=self._hex_to_rgb(s["bg_color"]), color=(0, 0, 0), width=0)

            # テキスト
            disp = self._create_sticky_display_text(
                s.get("title", ""), s.get("content", "")
            )
            shape.insert_textbox(
                fitz.Rect(rtx + 5, rty + 5, rtx + w - 5, rty + h - 5),
                disp,
                fontname="embed_font",
                fontsize=12,
                color=(0, 0, 0),
            )

        # 4. ノート (変更点: 全文テキストの描画)
        for key, info in self.notes_on_canvas.items():
            rtx, rty = tx(info["x"]), ty(info["y"])
            col = self.parent_app.key_colors.get(info["cp_key"].lower(), "#aaaaaa")

            shape.draw_rect(fitz.Rect(rtx, rty, rtx + 160, rty + 80))
            shape.finish(fill=self._hex_to_rgb(col), color=(0, 0, 0), width=1)

            full_text = f"[[{key}: {info['title']}]]"

            shape.insert_textbox(
                fitz.Rect(rtx + 5, rty + 5, rtx + 155, rty + 75),
                full_text,
                fontname="embed_font",
                fontsize=10,
                align=1,  # Center
                color=(0, 0, 0),
            )

        shape.commit()

        # --- 出力処理 (拡張子による分岐) ---
        suffix = output_path.suffix.lower()

        if suffix in [".png", ".jpg", ".jpeg"]:
            # 画像として保存 (dpi=150で十分高画質ですが、配布用なら300でも可)
            pix = page.get_pixmap(dpi=150)
            pix.save(str(output_path))
            doc.close()
            return

        else:
            # PDFとして保存 (正規化プロセスを経由)
            with tempfile.TemporaryDirectory() as temp_dir:
                temp_raw_pdf = Path(temp_dir) / "raw_canvas_export.pdf"
                doc.save(str(temp_raw_pdf))
                doc.close()

                # 5. 正規化とメタデータ埋め込み
                normalize_pdf_to_papersize(
                    str(temp_raw_pdf),
                    str(output_path),
                    595.276,
                    841.89,
                    target_format="A4",
                )
                embed_processing_flag(str(output_path))

                # 以下、メタデータ埋め込み処理
                key_rect_tuple = (0, 13, 391, 73)
                text_color = (0, 0, 0)
                if index_key:
                    hex_color = self.parent_app.key_colors.get(index_key.lower())
                    if hex_color:
                        text_color = self._hex_to_rgb(hex_color)

                cited_keys = sorted(list(self.notes_on_canvas.keys()))
                refs_qr_size_str = self._get_config_value(
                    "Extraction", "refs_qr_size", "75"
                )
                try:
                    refs_qr_size_pt = int(refs_qr_size_str)
                except ValueError:
                    refs_qr_size_pt = 75

                add_metadata_to_clip(
                    pdf_path_str=str(output_path),
                    font_path=str(font_path),
                    paper_width=595.276,
                    paper_height=841.89,
                    key_rect_tuple=key_rect_tuple,
                    index_key_to_embed=index_key,
                    text_color=text_color,
                    comment_to_embed=f"Canvas Export: {title}",
                    base_name=output_path.stem,
                    cited_keys_list=cited_keys,
                    refs_qr_size_pt=refs_qr_size_pt,
                    extra_keywords=["Synapsen:Whiteboard"],
                )

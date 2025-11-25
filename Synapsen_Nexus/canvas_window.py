import os
import re
import sys
import tempfile
import customtkinter as ctk
import tkinter as tk
from tkinter import filedialog, messagebox
import json
from pathlib import Path
import configparser
import fitz  # PyMuPDF
import sqlite3
import datetime
import shutil

# 分割したモジュールをインポート
import logging

from utils import _update_note_links

# Normalisiererモジュールのインポート設定
current_dir = Path(__file__).parent
normalisierer_dir = current_dir.parent / "Synapsen_Normalisierer"
if str(normalisierer_dir) not in sys.path:
    sys.path.append(str(normalisierer_dir))

try:
    from pdf_utils import (  # type: ignore
        convert_document_to_pdf,
        high_fidelity_flatten,
        normalize_pdf_to_papersize,
        add_metadata_to_clip,
        embed_processing_flag,
    )
except ImportError:
    pass

# ==============================================================================
# ロギング設定の初期化
# ==============================================================================
# 親ディレクトリ(ルート)をパスに追加して logging_setup.py をインポート可能にする
current_dir = Path(__file__).parent
root_dir = current_dir.parent
if str(root_dir) not in sys.path:
    sys.path.append(str(root_dir))

try:
    from logging_setup import setup_logging

    # アプリ名を指定して初期化
    setup_logging("Synapsen_Normalisierer")
    logger = logging.getLogger("Normalisierer")  # このファイル用のロガー取得
except ImportError:
    # logging_setup.py がない場合のフォールバック（print出力）
    print("Warning: logging_setup.py not found. Logging disabled.")

    class MockLogger:
        def info(self, msg):
            print(f"[INFO] {msg}")

        def error(self, msg, exc_info=None):
            print(f"[ERROR] {msg} {exc_info if exc_info else ''}")

        def warning(self, msg):
            print(f"[WARN] {msg}")

    logger = MockLogger()
# ==============================================================================


class StickyNoteDialog(ctk.CTkToplevel):
    def __init__(self, parent, title_val="", content_val="", color_val="#FFFFA5"):
        super().__init__(parent)
        self.title("付箋の編集")
        self.geometry("450x450")

        # アイコン設定 (遅延適用)
        self._custom_icon_path = None
        if hasattr(parent, "_custom_icon_path") and parent._custom_icon_path:
            self._custom_icon_path = parent._custom_icon_path
            self.after(200, lambda: self._apply_icon())

        self.grab_set()
        self.result = None
        self.selected_color = tk.StringVar(value=color_val)

        self.grid_columnconfigure(0, weight=1)
        self.grid_rowconfigure(5, weight=1)

        # --- 色選択 ---
        ctk.CTkLabel(self, text="色:", anchor="w").grid(
            row=0, column=0, padx=10, pady=(10, 2), sticky="ew"
        )

        color_frame = ctk.CTkFrame(self, fg_color="transparent")
        color_frame.grid(row=1, column=0, padx=10, pady=(0, 10), sticky="ew")

        colors = [
            ("イエロー", "#FFFFA5"),
            ("ブルー", "#D1EAFF"),
            ("レッド", "#FFD1D1"),
            ("グリーン", "#D1FFD1"),
            ("ホワイト", "#FFFFF0"),
            ("グレー", "#E0E0E0"),
        ]

        for i, (name, code) in enumerate(colors):

            if ctk.get_appearance_mode() == "Dark":
                text_color = "gray90"
            else:
                text_color = "black"

            btn = ctk.CTkRadioButton(
                color_frame,
                text=name,
                value=code,
                variable=self.selected_color,
                fg_color=code,
                border_color="gray",
                text_color=text_color,
                width=80,
            )
            r, c = divmod(i, 3)
            btn.grid(row=r, column=c, padx=5, pady=5, sticky="w")

        # --- テキスト入力 ---
        ctk.CTkLabel(self, text="タイトル (改行不可):", anchor="w").grid(
            row=2, column=0, padx=10, pady=(5, 2), sticky="ew"
        )
        self.title_entry = ctk.CTkEntry(self)
        self.title_entry.grid(row=3, column=0, padx=10, pady=(0, 10), sticky="ew")
        self.title_entry.insert(0, title_val)

        ctk.CTkLabel(self, text="内容:", anchor="w").grid(
            row=4, column=0, padx=10, pady=(0, 2), sticky="ew"
        )
        self.content_textbox = ctk.CTkTextbox(self)
        self.content_textbox.grid(row=5, column=0, padx=10, pady=(0, 10), sticky="nsew")
        self.content_textbox.insert("1.0", content_val)

        # ボタン
        btn_frame = ctk.CTkFrame(self, fg_color="transparent")
        btn_frame.grid(row=6, column=0, padx=10, pady=10, sticky="e")

        ctk.CTkButton(
            btn_frame,
            text="キャンセル",
            width=80,
            fg_color="gray",
            command=self.destroy,
        ).pack(side="left", padx=5)
        ctk.CTkButton(btn_frame, text="OK", width=80, command=self.on_ok).pack(
            side="left", padx=5
        )

    def _apply_icon(self):
        if self._custom_icon_path:
            try:
                self.iconbitmap(self._custom_icon_path)
            except Exception:
                pass

    def iconbitmap(self, *args, **kwargs):
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

    def on_ok(self):
        raw_title = self.title_entry.get().replace("\n", " ").replace("\r", "").strip()
        raw_content = self.content_textbox.get("1.0", "end-1c")
        color = self.selected_color.get()
        self.result = (raw_title, raw_content, color)
        self.destroy()


class ConversionDialog(ctk.CTkToplevel):
    """付箋をノートに変換する際に、タイトルとIndexKeyを指定するダイアログ"""

    def __init__(self, parent, default_title, key_options):
        super().__init__(parent)
        self.title("ノートに変換")
        self.geometry("400x250")

        self._custom_icon_path = None
        if hasattr(parent, "_custom_icon_path") and parent._custom_icon_path:
            self._custom_icon_path = parent._custom_icon_path
            self.after(200, lambda: self._apply_icon())

        self.grab_set()
        self.result = None
        self.key_options = key_options

        self.grid_columnconfigure(0, weight=1)

        ctk.CTkLabel(self, text="タイトル:", anchor="w").grid(
            row=0, column=0, padx=20, pady=(20, 5), sticky="ew"
        )
        self.title_entry = ctk.CTkEntry(self)
        self.title_entry.grid(row=1, column=0, padx=20, pady=(0, 15), sticky="ew")
        self.title_entry.insert(0, default_title)

        ctk.CTkLabel(self, text="Index Key (分類):", anchor="w").grid(
            row=2, column=0, padx=20, pady=(0, 5), sticky="ew"
        )
        self.key_combo = ctk.CTkComboBox(self, values=self.key_options)
        self.key_combo.grid(row=3, column=0, padx=20, pady=(0, 20), sticky="ew")
        if self.key_options:
            self.key_combo.set(self.key_options[0])

        btn_frame = ctk.CTkFrame(self, fg_color="transparent")
        btn_frame.grid(row=4, column=0, padx=20, pady=10, sticky="e")

        ctk.CTkButton(
            btn_frame,
            text="キャンセル",
            width=80,
            fg_color="gray",
            command=self.destroy,
        ).pack(side="left", padx=5)
        ctk.CTkButton(btn_frame, text="変換実行", width=80, command=self.on_ok).pack(
            side="left", padx=5
        )

    def _apply_icon(self):
        if self._custom_icon_path:
            try:
                self.iconbitmap(self._custom_icon_path)
            except Exception:
                pass

    def iconbitmap(self, *args, **kwargs):
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

    def on_ok(self):
        title = self.title_entry.get().strip()
        if not title:
            title = "NOTITLE"
        key = self.key_combo.get()
        self.result = (title, key)
        self.destroy()


class CanvasHelpWindow(ctk.CTkToplevel):
    def __init__(self, parent):
        super().__init__(parent)
        self.title("Canvas ヘルプ")
        self.geometry("600x650")

        if hasattr(parent, "_custom_icon_path") and parent._custom_icon_path:
            self._custom_icon_path = parent._custom_icon_path
            self.after(200, lambda: self._apply_icon())

        self.grab_set()

        help_text = """
■ 基本操作
---------------------------
[選択ツール]
・クリック: アイテムを選択
・ドラッグ: アイテムを移動
・Shift + クリック: 複数選択の追加/解除
・背景ドラッグ: 範囲選択 (ラバーバンド)

[視点操作]
・マウスホイール: 縦スクロール
・Shift + ホイール: 横スクロール
・Ctrl + ホイール: ズームイン/アウト
・ツールバーの +/- ボタン: ズーム操作

■ ツール
---------------------------
[付箋 (Sticky)]
・ドラッグまたはクリックで付箋を作成します。
・作成・編集ダイアログで、タイトル・内容・色を選択できます。
・ダブルクリック: 再編集
・サイズ変更:
    選択時に右下に表示される「■ハンドル」をドラッグすることで、サイズを自由に変更できます。
    ※ PDF出力時、付箋のサイズからはみ出した文字は印字されません。
       長文を入力した場合は、必ずサイズを広げて文字が収まるように調整してください。
・右クリック -> 「PDFを作成 (保存のみ)」:
    付箋の内容を正規化済みPDFとして出力します（Markdownファイルは残りません）。
    ※ 本文にMarkdown形式で記述すると、PDF出力時に見出しやリスト等が反映されます。
    ファイルは config.ini の nexus_output_folder に保存されます。
    ★作成時にIndex Keyを選択し、QRコードとして1ページ目に埋め込みます。

[⇔ 接続] (コネクタ)
・ノート/付箋から別のノート/付箋へドラッグして接続します。
・ダブルクリック:
    - 未リンク (白/黒): DBに相互リンクを追記し、緑色にします。
    - リンク済 (緑色): DBからリンクを削除し、元の色に戻します。
・右クリック -> 「リンク作成」「リンク解除」も可能です。

[□ 枠 / ／ 線]
・装飾用の図形を描画します。

■ 連携機能
---------------------------
[OR検索]
・選択したノートのKeyで `[[Key1]] OR [[Key2]]...` 検索を実行します。

[ノート追加]
・Nexus本体で選択中のノートをキャンバスに追加します。

[PDF出力]
・キャンバス全体をPDFとして保存します。
・ノート、付箋、接続線、枠線などがそのまま出力されます。
"""
        textbox = ctk.CTkTextbox(self, wrap="word", font=("", 12))
        textbox.pack(fill="both", expand=True, padx=10, pady=10)
        textbox.insert("1.0", help_text)
        textbox.configure(state="disabled")

    def _apply_icon(self):
        if self._custom_icon_path:
            try:
                self.iconbitmap(self._custom_icon_path)
            except Exception:
                pass

    def iconbitmap(self, *args, **kwargs):
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


class CanvasWindow(ctk.CTkToplevel):
    def __init__(self, parent_app):
        super().__init__(parent_app)
        self.parent_app = parent_app
        self.title("Synapsen Canvas")
        self.geometry("1200x800")

        # --- アイコン設定 (初期化) ---
        self._custom_icon_path = None
        if hasattr(parent_app, "icon_path") and parent_app.icon_path:
            self._custom_icon_path = str(parent_app.icon_path)
            self.after(200, lambda: self._apply_icon())

        self.default_save_file = parent_app.base_path / "canvas_data.json"

        # --- データ構造 ---
        self.notes_on_canvas = {}
        self.stickies_on_canvas = []
        self.shapes_on_canvas = []
        self.connections_on_canvas = []

        self.selected_items = set()
        self.current_scale = 1.0
        self.base_font_size_note = 10
        self.base_font_size_sticky = 12
        self.base_line_width = 2

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

        # --- UI ---
        self.toolbar = ctk.CTkFrame(self)
        self.toolbar.pack(side="top", fill="x", padx=5, pady=5)

        ctk.CTkButton(
            self.toolbar,
            text="開く",
            width=50,
            command=self.load_from_file,
            fg_color="#585a9c",
        ).pack(side="left", padx=2)
        ctk.CTkButton(
            self.toolbar, text="保存", width=50, command=self.save_canvas
        ).pack(side="left", padx=2)
        ctk.CTkButton(
            self.toolbar, text="別名保存", width=60, command=self.save_as_file
        ).pack(side="left", padx=2)
        ctk.CTkLabel(self.toolbar, text="|").pack(side="left", padx=5)
        ctk.CTkButton(
            self.toolbar,
            text="PDF出力",
            width=60,
            command=self.export_canvas_dialog,
            fg_color="#00695C",
            hover_color="#004D40",
        ).pack(side="left", padx=2)

        # ズーム関連
        ctk.CTkLabel(self.toolbar, text="| ズーム:").pack(side="left", padx=5)
        # [-] ボタン
        ctk.CTkButton(
            self.toolbar, text="-", width=30, command=self._zoom_out_btn
        ).pack(side="left", padx=2)

        # リセットボタン
        self.zoom_label_var = ctk.StringVar(value="100%")
        self.zoom_reset_button = ctk.CTkButton(
            self.toolbar,
            textvariable=self.zoom_label_var,  # テキストは動的に変わる
            width=50,
            command=self.reset_view,           # クリック時の動作
            font=("", 12),
            fg_color="transparent",            # 通常は背景透明（ラベル風）
            border_width=1,                    # 枠線をつけてボタンであることを示す
            border_color="gray",
            text_color=("black", "white"),
            hover_color=("gray70", "gray30"),  # ホバー時の色
        )
        self.zoom_reset_button.pack(side="left", padx=2)

        # [+] ボタン
        ctk.CTkButton(self.toolbar, text="+", width=30, command=self._zoom_in_btn).pack(
            side="left", padx=2
        )

        ctk.CTkLabel(self.toolbar, text="| ツール:").pack(side="left", padx=5)

        self.mode_buttons = {}
        self.status_label_var = ctk.StringVar(value="現在のツール: 選択/移動")
        self.status_label = ctk.CTkLabel(
            self.toolbar,
            textvariable=self.status_label_var,
            font=("", 12, "bold"),
            text_color="gray",
        )

        def set_mode(mode, label_text):
            self.current_mode = mode
            self.status_label_var.set(f"現在のツール: {label_text}")
            for m, btn in self.mode_buttons.items():
                btn.configure(fg_color="#1F6AA5" if m == mode else "#777777")
            if mode != "select":
                self._clear_selection()

        self.mode_buttons["select"] = ctk.CTkButton(
            self.toolbar,
            text="選択",
            width=50,
            command=lambda: set_mode("select", "選択/移動"),
        )
        self.mode_buttons["select"].pack(side="left", padx=2)
        self.mode_buttons["connect"] = ctk.CTkButton(
            self.toolbar,
            text="⇔ 接続",
            width=50,
            command=lambda: set_mode("connect", "ノート接続 (リンク)"),
        )
        self.mode_buttons["connect"].pack(side="left", padx=2)
        self.mode_buttons["sticky"] = ctk.CTkButton(
            self.toolbar,
            text="付箋",
            width=50,
            command=lambda: set_mode("sticky", "付箋作成"),
        )
        self.mode_buttons["sticky"].pack(side="left", padx=2)
        self.mode_buttons["rect"] = ctk.CTkButton(
            self.toolbar,
            text="□ 枠",
            width=50,
            command=lambda: set_mode("rect", "枠線"),
        )
        self.mode_buttons["rect"].pack(side="left", padx=2)
        self.mode_buttons["line"] = ctk.CTkButton(
            self.toolbar,
            text="／ 線",
            width=50,
            command=lambda: set_mode("line", "線 (Path)"),
        )
        self.mode_buttons["line"].pack(side="left", padx=2)

        ctk.CTkLabel(self.toolbar, text="|").pack(side="left", padx=5)
        ctk.CTkButton(
            self.toolbar,
            text="OR検索",
            width=70,
            command=self.send_or_search_to_nexus,
            fg_color="#E0a800",
            hover_color="#D09800",
            text_color="black",
        ).pack(side="left", padx=2)

        # 右側ボタン
        ctk.CTkButton(
            self.toolbar,
            text="全消去",
            width=60,
            fg_color="#D9534F",
            hover_color="#C9302C",
            command=self.clear_canvas,
        ).pack(side="right", padx=5)
        ctk.CTkButton(
            self.toolbar, text="ノート追加", command=self.add_selected_notes
        ).pack(side="right", padx=5)
        ctk.CTkButton(self.toolbar, text="？", width=30, command=self.show_help).pack(
            side="right", padx=5
        )

        self.status_label.pack(side="right", padx=15)
        set_mode("select", "選択/移動")

        self.canvas_frame = ctk.CTkFrame(self)
        self.canvas_frame.pack(fill="both", expand=True, padx=5, pady=5)

        if ctk.get_appearance_mode() == "Dark":
            self.bg_color = "#2b2b2b"
        else:
            self.bg_color = "#f0f0f0"
        if ctk.get_appearance_mode() == "Dark":
            self.grid_color = "#3a3a3a"
        else:
            self.grid_color = "#e0e0e0"

        self.canvas = tk.Canvas(
            self.canvas_frame, bg=self.bg_color, highlightthickness=0
        )
        hbar = tk.Scrollbar(
            self.canvas_frame, orient=tk.HORIZONTAL, command=self.canvas.xview
        )
        hbar.pack(side=tk.BOTTOM, fill=tk.X)
        vbar = tk.Scrollbar(
            self.canvas_frame, orient=tk.VERTICAL, command=self.canvas.yview
        )
        vbar.pack(side=tk.RIGHT, fill=tk.Y)
        self.canvas.config(xscrollcommand=hbar.set, yscrollcommand=vbar.set)
        self.canvas.pack(fill="both", expand=True)

        self._draw_grid(width=5000, height=5000)

        self.canvas.bind("<ButtonPress-1>", self.on_canvas_click)
        self.canvas.bind("<B1-Motion>", self.on_canvas_drag)
        self.canvas.bind("<ButtonRelease-1>", self.on_canvas_release)
        self.canvas.bind("<Double-Button-1>", self.on_double_click)
        self.canvas.bind("<Button-3>", self.on_right_click)
        self.canvas.bind("<MouseWheel>", self._on_mousewheel)
        self.canvas.bind("<Shift-MouseWheel>", self._on_shift_mousewheel)
        self.canvas.bind("<Control-MouseWheel>", self._on_zoom_wheel)

        self.load_canvas_data(self.default_save_file)

        self.lift()
        self.attributes("-topmost", True)

        # window描写待ちの遅延が必要な処理
        self.after(500, lambda: self.attributes("-topmost", False))
        self.after(100, self.center_view)  # 初期表示位置(中央)に移動
        self.focus_force()

    def show_help(self):
        if hasattr(self, "help_window") and self.help_window.winfo_exists():
            self.help_window.focus()
            return
        self.help_window = CanvasHelpWindow(self)
        self.help_window.focus()

    def _apply_icon(self):
        if self._custom_icon_path:
            try:
                self.iconbitmap(self._custom_icon_path)
            except Exception:
                pass

    def iconbitmap(self, *args, **kwargs):
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

    def _draw_grid(self, width, height, step=50):
        self.canvas.delete("grid")
        scaled_step = step * self.current_scale

        # --- 1. 通常のグリッド描画 ---
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

        # --- 2. 用紙サイズ補助枠 & 中央十字 の描画 ---
        base_pw = getattr(self.parent_app, "paper_width", 595.276)
        base_ph = getattr(self.parent_app, "paper_height", 841.89)
        p_size_name = self.parent_app.config_data.get("paper_size", "A4")

        # サイズを4倍に変更
        pw = base_pw * 4
        ph = base_ph * 4

        # キャンバスの仮想中心 (論理座標)
        cx = width / 2
        cy = height / 2

        # 枠の座標計算
        x1 = (cx - pw / 2) * self.current_scale
        y1 = (cy - ph / 2) * self.current_scale
        x2 = (cx + pw / 2) * self.current_scale
        y2 = (cy + ph / 2) * self.current_scale

        # スタイル設定
        guide_color = "#FF4081" if ctk.get_appearance_mode() == "Dark" else "#D81B60"

        # 枠線の描画
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

        # ラベル
        self.canvas.create_text(
            x1 + 10,
            y1 + 10,
            text=f"{p_size_name} (x4) ガイド",
            anchor="nw",
            fill=guide_color,
            font=("", int(12 * self.current_scale)),
            tags=("grid", "guide"),
        )

        # 中央十字 (Crosshair) の描画
        cross_size = 20 * self.current_scale
        center_x = cx * self.current_scale
        center_y = cy * self.current_scale

        # 横線
        self.canvas.create_line(
            center_x - cross_size,
            center_y,
            center_x + cross_size,
            center_y,
            fill=guide_color,
            width=2,
            tags=("grid", "guide"),
        )
        # 縦線
        self.canvas.create_line(
            center_x,
            center_y - cross_size,
            center_x,
            center_y + cross_size,
            fill=guide_color,
            width=2,
            tags=("grid", "guide"),
        )

        # 最背面へ移動
        self.canvas.tag_lower("grid")

        # スクロール領域の更新
        self.canvas.configure(scrollregion=self.canvas.bbox("all"))

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
        """キャンバスの中央が画面の中心に来るようにスクロールする"""
        self.update_idletasks()  # ウィンドウサイズを確定させる

        try:
            # 現在のスクロール領域を取得 (例: "0 0 5000 5000")
            sr = self.canvas.cget("scrollregion").split()
            if not sr:
                return

            # floatに変換
            x1, y1, x2, y2 = map(float, sr)
            content_w = x2 - x1
            content_h = y2 - y1

            if content_w <= 0 or content_h <= 0:
                return

            # ビューポート（表示領域）のサイズ
            view_w = self.canvas.winfo_width()
            view_h = self.canvas.winfo_height()

            # 中央を表示するためのオフセット比率を計算
            # (コンテンツ全体の中心 - 表示領域の半分) / コンテンツ全体幅
            x_fraction = ((content_w / 2) - (view_w / 2)) / content_w
            y_fraction = ((content_h / 2) - (view_h / 2)) / content_h

            self.canvas.xview_moveto(x_fraction)
            self.canvas.yview_moveto(y_fraction)

        except Exception as e:
            print(f"Centering failed: {e}")

    def reset_view(self):
        """
        ズーム倍率を切り替え、中央にスクロールする。
        - 現在が 100% (約1.0) の場合 -> 30% (0.3) に変更 (A4全体確認用)
        - それ以外の場合 -> 100% (1.0) に戻す (標準)
        """
        if self.current_scale == 0:
            return

        target_scale = 1.0

        # 現在がほぼ100% (誤差許容) なら、ターゲットを30%に設定
        if 0.99 < self.current_scale < 1.01:
            target_scale = 0.3

        # 現在の倍率からターゲット倍率へ移行するための係数を計算
        # (例: 現在1.0 -> 目標0.3 なら、係数は 0.3)
        scale_factor = target_scale / self.current_scale

        # ズームを実行 (内部で current_scale 更新とラベル書き換えが行われる)
        self.zoom(scale_factor)

        # 中央へ移動
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

        # ラベルの更新
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

        for key, info in self.notes_on_canvas.items():
            rect_id, text_id = info["ids"]
            self.canvas.itemconfigure(text_id, font=new_note_font)
            is_sel = ("note", key) in self.selected_items
            self.canvas.itemconfigure(
                rect_id,
                width=selected_width if is_sel else base_width,
                outline="#585a9c" if is_sel else "white",
            )

        for sticky in self.stickies_on_canvas:
            rect_id, text_id = sticky["ids"]
            self.canvas.itemconfigure(text_id, font=new_sticky_font)
            is_sel = ("sticky", str(id(sticky))) in self.selected_items
            self.canvas.itemconfigure(
                rect_id,
                width=selected_width if is_sel else 0,
                outline="#585a9c" if is_sel else "",
            )

        for s in self.shapes_on_canvas:
            if s["type"] in ("rect", "line"):
                is_sel = ("shape", str(id(s))) in self.selected_items
                w = selected_width if is_sel else base_width
                if s["type"] == "rect":
                    c = "#585a9c" if is_sel else "red"
                    self.canvas.itemconfigure(s["id"], outline=c, width=w)
                else:
                    c = (
                        "#585a9c"
                        if is_sel
                        else ("white" if self.bg_color == "#2b2b2b" else "black")
                    )
                    self.canvas.itemconfigure(s["id"], fill=c, width=w)

        for c in self.connections_on_canvas:
            self.canvas.itemconfigure(c["id"], width=base_width)

    # --- Selection ---
    def _get_item_key(self, item_type, data):
        if isinstance(data, str):
            return data
        if item_type == "note":
            return data
        if item_type == "sticky":
            return str(id(data))
        if item_type == "shape":
            return str(id(data))
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
        # まず既存のリサイズハンドルを全消去
        self.canvas.delete("resize_handle")

        base_width = max(1, int(self.base_line_width * self.current_scale))
        selected_width = max(2, int(4 * self.current_scale))
        sel_col = "#585a9c"

        # ハンドルのサイズ
        handle_size = 10 * self.current_scale

        # ノートの処理 (既存のまま)
        for key, info in self.notes_on_canvas.items():
            # ... (省略: 既存のコード) ...
            rect_id = info["ids"][0]
            if ("note", key) in self.selected_items:
                self.canvas.itemconfigure(
                    rect_id, outline=sel_col, width=selected_width
                )
                self.canvas.tag_raise(rect_id)
                self.canvas.tag_raise(info["ids"][1])
            else:
                self.canvas.itemconfigure(rect_id, outline="white", width=base_width)

        # 付箋の処理 (ハンドル描画を追加)
        for sticky in self.stickies_on_canvas:
            rect_id = sticky["ids"][0]
            sticky_key = str(id(sticky))

            if ("sticky", sticky_key) in self.selected_items:
                self.canvas.itemconfigure(
                    rect_id, outline=sel_col, width=selected_width
                )
                self.canvas.tag_raise(rect_id)
                self.canvas.tag_raise(sticky["ids"][1])

                # --- 追加: リサイズハンドルの描画 ---
                # 付箋の右下座標を計算
                x1, y1, x2, y2 = self.canvas.coords(rect_id)

                # ハンドル (四角形) を作成
                hid = self.canvas.create_rectangle(  # noqa: F841
                    x2 - handle_size,
                    y2 - handle_size,
                    x2,
                    y2,
                    fill="gray",
                    outline="white",
                    tags=("resize_handle", f"handle_{sticky_key}"),
                )
                # sticky_key を紐付けておくことで、クリック時に親を特定できる

            else:
                self.canvas.itemconfigure(rect_id, outline="", width=0)

        for s in self.shapes_on_canvas:
            if s["type"] in ("rect", "line"):
                is_sel = ("shape", str(id(s))) in self.selected_items
                w = selected_width if is_sel else base_width
                if s["type"] == "rect":
                    c = sel_col if is_sel else "red"
                    self.canvas.itemconfigure(s["id"], outline=c, width=w)
                else:
                    c = (
                        sel_col
                        if is_sel
                        else ("white" if self.bg_color == "#2b2b2b" else "black")
                    )
                    self.canvas.itemconfigure(s["id"], fill=c, width=w)

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

    # --- Creation ---
    def add_selected_notes(self):
        selected_keys = self.parent_app.selected_keys
        if not selected_keys:
            messagebox.showinfo(
                "情報", "Nexus本体でノートを選択してください。", parent=self
            )
            return
        df = self.parent_app.df
        start_x = self.canvas.canvasx(100)
        start_y = self.canvas.canvasy(100)
        added = 0
        for _, row in df[df["key"].isin(selected_keys)].iterrows():
            key = row["key"]
            if key in self.notes_on_canvas:
                continue
            self.create_note_item(
                key, row["title"], row.get("commonplace_key", ""), start_x, start_y
            )
            start_x += 30 * self.current_scale
            start_y += 30 * self.current_scale
            added += 1
        if added > 0:
            self.save_canvas()

    def create_note_item(self, key, title, cp_key, x, y):
        # x, y は論理座標(スケール1.0)として受け取る
        color = self.parent_app.key_colors.get(cp_key.lower(), "#aaaaaa")

        # 描画用サイズ・座標計算 (画面表示用)
        sw, sh = 160 * self.current_scale, 80 * self.current_scale
        sx, sy = x * self.current_scale, y * self.current_scale

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
        dt = (title[:20] + "..") if len(title) > 20 else title
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

        # 内部データには論理座標 (x, y) をそのまま保存
        self.notes_on_canvas[key] = {
            "x": x,
            "y": y,
            "title": title,
            "cp_key": cp_key,
            "ids": (rid, tid),
        }

    def _create_sticky_display_text(self, title, content):
        """
        タイトルと内容を結合して表示用テキストを作成する。
        行数制限を行わず、常に全文を返し、表示の調整は不戦のリサイズで行ってもらう。
        """
        return f"{title}\n{content}"

    def create_sticky_item(
        self, x, y, title="", content="", bg_color="#FFFFA5", w=180, h=120
    ):
        # 描画用サイズ・座標計算
        sw, sh = w * self.current_scale, h * self.current_scale
        sx, sy = x * self.current_scale, y * self.current_scale

        # パディングもスケールさせる (論理10px)
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
            "x": x,
            "y": y,
            "w": w,
            "h": h,  # 論理座標を保存
            "title": title,
            "content": content,
            "bg_color": bg_color,
            "ids": (rid, tid),
            "shadow_id": sid,
        }

        real_tag = f"sticky_{id(item)}"
        self.canvas.addtag_withtag(real_tag, rid)
        self.canvas.addtag_withtag(real_tag, tid)
        self.canvas.addtag_withtag(real_tag, sid)

        self.stickies_on_canvas.append(item)
        return item

    def create_connection_item(self, from_item, to_item):
        f_type, f_key = from_item
        t_type, t_key = to_item
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

        if is_linked:
            col = "#28a745"
        else:
            if self.bg_color == "#2b2b2b":
                col = "white"
            else:
                col = "black"

        lw = max(1, int(self.base_line_width * self.current_scale))

        lid = self.canvas.create_line(
            x1, y1, x2, y2, fill=col, width=lw, arrow=tk.LAST, tags=("connection",)
        )
        try:
            self.canvas.tag_lower(lid, "note")
        except tk.TclError:
            pass  # ノートがない場合は何もしない

        try:
            self.canvas.tag_lower(lid, "sticky")
        except tk.TclError:
            pass  # 付箋がない場合は何もしない

        try:
            self.canvas.tag_raise(lid, "grid")
        except tk.TclError:
            # グリッドは通常存在しますが、念のため安全策
            self.canvas.tag_lower(lid)

        self.connections_on_canvas.append(
            {
                "id": lid,
                "from_key": f_key,
                "to_key": t_key,
                "from_type": f_type,
                "to_type": t_type,
            }
        )

    def _get_connection_coords(self, f_type, f_key, t_type, t_key):
        def get_bbox(type_, key_):
            if type_ == "note":
                if key_ in self.notes_on_canvas:
                    return self.canvas.bbox(self.notes_on_canvas[key_]["ids"][0])
            elif type_ == "sticky":
                for s in self.stickies_on_canvas:
                    if str(id(s)) == key_:
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

    def create_shape_item(self, shape_type, x=0, y=0, w=0, h=0, text="", x2=0, y2=0):
        lw = max(1, int(self.base_line_width * self.current_scale))

        # 描画用座標変換
        sx, sy = x * self.current_scale, y * self.current_scale
        sw, sh = w * self.current_scale, h * self.current_scale
        sx2, sy2 = x2 * self.current_scale, y2 * self.current_scale

        if shape_type == "rect":
            iid = self.canvas.create_rectangle(
                sx,
                sy,
                sx + sw,
                sy + sh,
                outline="red",
                width=lw,
                dash=(4, 4),
                tags=("shape", "rect"),
            )
            self.shapes_on_canvas.append(
                {"id": iid, "type": "rect", "x": x, "y": y, "w": w, "h": h}  # 論理座標
            )
            self.canvas.addtag_withtag(f"shape_{id(self.shapes_on_canvas[-1])}", iid)

        elif shape_type == "line":
            col = "white" if self.bg_color == "#2b2b2b" else "black"
            iid = self.canvas.create_line(
                sx, sy, sx2, sy2, fill=col, width=lw, tags=("shape", "line")
            )
            self.shapes_on_canvas.append(
                {
                    "id": iid,
                    "type": "line",
                    "x1": x,
                    "y1": y,
                    "x2": x2,
                    "y2": y2,  # 論理座標
                }
            )
            self.canvas.addtag_withtag(f"shape_{id(self.shapes_on_canvas[-1])}", iid)

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

    # --- Event Handlers ---
    def on_canvas_click(self, event):
        cx, cy = self.canvas.canvasx(event.x), self.canvas.canvasy(event.y)

        # 【重要】ドラッグ移動量計算用 (直前の座標)
        self.drag_data["x"] = cx
        self.drag_data["y"] = cy

        # 【重要】ラバーバンド/シェイプ生成の始点用
        self.drag_data["start_x"] = cx
        self.drag_data["start_y"] = cy

        self.drag_data["moved"] = False

        # --- リサイズ判定やアイテム選択ロジック ---
        clicked_items = self.canvas.find_overlapping(cx - 1, cy - 1, cx + 1, cy + 1)
        for item in clicked_items:
            tags = self.canvas.gettags(item)
            if "resize_handle" in tags:
                # タグから付箋IDを特定 (例: "handle_12345")
                for t in tags:
                    if t.startswith("handle_"):
                        sticky_id = t.split("_")[1]
                        self.drag_data["start_item"] = ("sticky_resize", sticky_id)
                        self.current_mode = "resize_sticky"  # 一時的にモード変更
                        return

        items = self.canvas.find_overlapping(cx - 1, cy - 1, cx + 1, cy + 1)
        target = None

        for item in reversed(items):
            tags = self.canvas.gettags(item)
            if "note" in tags:
                for t in tags:
                    if t not in ("note", "current"):
                        target = ("note", t)
                        break
            elif "sticky" in tags:
                for s in self.stickies_on_canvas:
                    if s["ids"][0] == item or s["ids"][1] == item:
                        target = ("sticky", str(id(s)))
                        break
            elif "shape" in tags:
                for s in self.shapes_on_canvas:
                    if s["id"] == item:
                        target = ("shape", str(id(s)))
                        break
            if target:
                break

        is_shift = (event.state & 0x0001) != 0

        if self.current_mode == "connect":
            if target and target[0] in ("note", "sticky"):
                self.drag_data["start_item"] = target
                col = "white" if self.bg_color == "#2b2b2b" else "black"
                self.drag_data["temp_id"] = self.canvas.create_line(
                    cx, cy, cx, cy, fill=col, width=2, dash=(2, 2)
                )
            return

        if self.current_mode == "select":
            if target:
                if is_shift:
                    self._toggle_selection(*target)
                else:
                    if target not in self.selected_items:
                        self._select_item(*target)

                self.drag_data["start_item"] = target
                self.drag_data["x"], self.drag_data["y"] = cx, cy
            else:
                if not is_shift:
                    self._clear_selection()
                self.drag_data["rubberband_id"] = self.canvas.create_rectangle(
                    cx, cy, cx, cy, outline="#585a9c", dash=(2, 2)
                )

        elif self.current_mode == "sticky":
            dialog = StickyNoteDialog(self)
            self.wait_window(dialog)
            if dialog.result:
                title, content, color = dialog.result
                if title or content:
                    self.create_sticky_item(cx, cy, title, content, bg_color=color)
                    self.save_canvas()

        elif self.current_mode in ("rect", "line"):
            if self.current_mode == "rect":
                self.drag_data["temp_id"] = self.canvas.create_rectangle(
                    cx, cy, cx, cy, outline="red"
                )
            elif self.current_mode == "line":
                self.drag_data["temp_id"] = self.canvas.create_line(
                    cx, cy, cx, cy, fill="white"
                )

    def on_canvas_drag(self, event):
        cx, cy = self.canvas.canvasx(event.x), self.canvas.canvasy(event.y)

        # 差分計算
        dx = cx - self.drag_data["x"]
        dy = cy - self.drag_data["y"]

        self.drag_data["moved"] = True

        # 論理座標系への変換 (移動用)
        ldx = dx / self.current_scale
        ldy = dy / self.current_scale

        # 1. リサイズ処理 (ここを全面的に修正)
        if self.current_mode == "resize_sticky":
            target_type, target_key = self.drag_data["start_item"]
            if target_type == "sticky_resize":
                sticky_iter = (
                    s for s in self.stickies_on_canvas if str(id(s)) == target_key
                )
                target_sticky = next(sticky_iter, None)
                if target_sticky:

                    # 1. 現在の描画座標（左上）を取得
                    # これにより、ズームやパンで移動していても「見た目通りの位置」が基準になります
                    rect_id = target_sticky["ids"][0]
                    curr_coords = self.canvas.coords(rect_id)  # [x1, y1, x2, y2]
                    x1, y1 = curr_coords[0], curr_coords[1]

                    # 2. 現在の見た目のサイズを取得
                    current_screen_w = curr_coords[2] - x1
                    current_screen_h = curr_coords[3] - y1

                    # 3. マウス移動量(dx, dy)を加算して新しいサイズを決定
                    # (最小サイズ制限: 論理50px相当)
                    min_screen_size = 50 * self.current_scale
                    new_screen_w = max(min_screen_size, current_screen_w + dx)
                    new_screen_h = max(min_screen_size, current_screen_h + dy)

                    # 4. 内部データ(論理サイズ)を更新
                    # (見た目のサイズをスケールで割って論理値に戻して保存)
                    target_sticky["w"] = new_screen_w / self.current_scale
                    target_sticky["h"] = new_screen_h / self.current_scale

                    # 5. 描画の更新
                    # 5-1. 本体の更新
                    self.canvas.coords(
                        rect_id, x1, y1, x1 + new_screen_w, y1 + new_screen_h
                    )

                    # 5-2. 影の更新
                    self.canvas.coords(
                        target_sticky["shadow_id"],
                        x1 + 5,
                        y1 + 5,
                        x1 + new_screen_w + 5,
                        y1 + new_screen_h + 5,
                    )

                    # 5-3. テキスト (パディング計算含む)
                    pad = 10 * self.current_scale
                    self.canvas.coords(target_sticky["ids"][1], x1 + pad, y1 + pad)
                    self.canvas.itemconfigure(
                        target_sticky["ids"][1], width=new_screen_w - (pad * 2)
                    )

                    # 5-4. ハンドル
                    # 現在のハンドルIDを探す
                    handle_tag = f"handle_{target_key}"
                    handle_id = self.canvas.find_withtag(handle_tag)
                    if handle_id:
                        hs = 10 * self.current_scale
                        self.canvas.coords(
                            handle_id,
                            x1 + new_screen_w - hs,
                            y1 + new_screen_h - hs,
                            x1 + new_screen_w,
                            y1 + new_screen_h,
                        )

            # リサイズ時はマウス位置を更新
            self.drag_data["x"] = cx
            self.drag_data["y"] = cy
            return

        # 2. 接続線の仮描画
        if self.current_mode == "connect":
            if self.drag_data["temp_id"]:
                # 始点は start_x (クリック位置) を使用
                self.canvas.coords(
                    self.drag_data["temp_id"],
                    self.drag_data["start_x"],
                    self.drag_data["start_y"],
                    cx,
                    cy,
                )
            return

        # 3. 移動処理
        if self.current_mode == "select":
            if self.selected_items:
                for t, k in self.selected_items:
                    # 見た目の移動 (スクリーン座標 dx, dy を使用)
                    # 内部データの更新 (論理座標 ldx, ldy を使用)

                    if t == "note":
                        info = self.notes_on_canvas[k]
                        self.canvas.move(info["ids"][0], dx, dy)
                        self.canvas.move(info["ids"][1], dx, dy)
                        info["x"] += ldx
                        info["y"] += ldy
                        self.update_connections(t, k)
                    elif t == "sticky":
                        s = next(
                            (x for x in self.stickies_on_canvas if str(id(x)) == k),
                            None,
                        )
                        if s:
                            self.canvas.move(s["ids"][0], dx, dy)
                            self.canvas.move(s["ids"][1], dx, dy)
                            self.canvas.move(s["shadow_id"], dx, dy)
                            s["x"] += ldx
                            s["y"] += ldy
                            self.update_connections(t, k)

                            hid = self.canvas.find_withtag(f"handle_{k}")
                            if hid:
                                self.canvas.move(hid, dx, dy)

                    elif t == "shape":
                        s = next(
                            (x for x in self.shapes_on_canvas if str(id(x)) == k), None
                        )
                        if s:
                            self.canvas.move(s["id"], dx, dy)
                            if s["type"] == "line":
                                s["x1"] += ldx
                                s["y1"] += ldy
                                s["x2"] += ldx
                                s["y2"] += ldy
                            else:
                                s["x"] += ldx
                                s["y"] += ldy

                # 移動完了後に現在位置を保存
                self.drag_data["x"] = cx
                self.drag_data["y"] = cy

            elif self.drag_data["rubberband_id"]:
                # ラバーバンドは start_x (クリック位置) を基準にする
                self.canvas.coords(
                    self.drag_data["rubberband_id"],
                    self.drag_data["start_x"],
                    self.drag_data["start_y"],
                    cx,
                    cy,
                )

        # 4. シェイプ作成
        elif self.current_mode in ("rect", "line"):
            if self.drag_data["temp_id"]:
                self.canvas.coords(
                    self.drag_data["temp_id"],
                    self.drag_data["start_x"],
                    self.drag_data["start_y"],
                    cx,
                    cy,
                )

    def on_canvas_release(self, event):
        if self.current_mode == "resize_sticky":
            self.current_mode = "select"  # モードを戻す
            self.drag_data["start_item"] = None
            self.save_canvas()  # 保存
            self._update_selection_visuals()  # ハンドル位置などを綺麗に再描画
            return

        cx, cy = self.canvas.canvasx(event.x), self.canvas.canvasy(event.y)

        if self.current_mode == "select" and self.drag_data["rubberband_id"]:
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
            return

        if self.current_mode == "select" and self.drag_data["moved"]:
            self.save_canvas()

        if self.current_mode == "connect":
            if self.drag_data["temp_id"]:
                self.canvas.delete(self.drag_data["temp_id"])
                self.drag_data["temp_id"] = None
            if self.drag_data["start_item"]:
                items = self.canvas.find_overlapping(cx - 1, cy - 1, cx + 1, cy + 1)
                end_target = None
                for item in reversed(items):
                    tags = self.canvas.gettags(item)
                    if "note" in tags:
                        for t in tags:
                            if t not in ("note", "current"):
                                end_target = ("note", t)
                                break
                    elif "sticky" in tags:
                        for s in self.stickies_on_canvas:
                            if s["ids"][0] == item or s["ids"][1] == item:
                                end_target = ("sticky", str(id(s)))
                                break
                    if end_target:
                        break

                if end_target and end_target != self.drag_data["start_item"]:
                    self.create_connection_item(
                        self.drag_data["start_item"], end_target
                    )
                    self.save_canvas()
            self.drag_data["start_item"] = None
            return

        if self.current_mode in ("rect", "line") and self.drag_data["temp_id"]:
            self.canvas.delete(self.drag_data["temp_id"])
            if abs(cx - self.drag_data["start_x"]) > 5:
                # 作成時の座標を論理座標に変換して渡す
                lx1 = self.drag_data["start_x"] / self.current_scale
                ly1 = self.drag_data["start_y"] / self.current_scale
                lx2 = cx / self.current_scale
                ly2 = cy / self.current_scale

                self.create_shape_item(
                    self.current_mode,
                    x=lx1,
                    y=ly1,
                    x2=lx2,
                    y2=ly2,
                    w=abs(lx2 - lx1),
                    h=abs(ly2 - ly1),
                )
                self.save_canvas()

        self.drag_data["start_item"] = None

    def on_double_click(self, event):
        cx, cy = self.canvas.canvasx(event.x), self.canvas.canvasy(event.y)
        item = self.canvas.find_closest(cx, cy)[0]
        tags = self.canvas.gettags(item)

        if "sticky" in tags:
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
                    t, c, col = dialog.result
                    target["title"] = t
                    target["content"] = c
                    target["bg_color"] = col

                    display_text = self._create_sticky_display_text(t, c)
                    self.canvas.itemconfigure(target["ids"][1], text=display_text)
                    self.canvas.itemconfigure(target["ids"][0], fill=col)
                    self.save_canvas()

        elif "note" in tags:
            for t in tags:
                if t not in ("note", "current"):
                    self.parent_app.open_preview_window(t, ui_master=self)
                    break
        elif "connection" in tags:
            target_conn = None
            for c in self.connections_on_canvas:
                if c["id"] == item:
                    target_conn = c
                    break
            if (
                target_conn
                and target_conn["from_type"] == "note"
                and target_conn["to_type"] == "note"
            ):
                self._handle_connection_double_click(target_conn)

    def on_right_click(self, event):
        cx, cy = self.canvas.canvasx(event.x), self.canvas.canvasy(event.y)
        item = self.canvas.find_closest(cx, cy)[0]
        tags = self.canvas.gettags(item)

        target = None
        if "note" in tags:
            for t in tags:
                if t not in ("note", "current"):
                    target = ("note", t)
                    break
        elif "sticky" in tags:
            for s in self.stickies_on_canvas:
                if s["ids"][0] == item or s["ids"][1] == item:
                    target = ("sticky", s)
                    break
        elif "shape" in tags:
            for s in self.shapes_on_canvas:
                if s["id"] == item:
                    target = ("shape", s)
                    break
        elif "connection" in tags:
            for c in self.connections_on_canvas:
                if c["id"] == item:
                    target = ("connection", c)
                    break

        if target:
            menu = tk.Menu(self, tearoff=0)
            type_, obj = target

            if type_ == "sticky":
                menu.add_command(
                    label="PDFを作成 (保存のみ)",
                    command=lambda: self._convert_sticky_to_note_pipeline(obj),
                )
                selected_stickies = [
                    i[1] for i in self.selected_items if i[0] == "sticky"
                ]
                is_obj_selected = str(id(obj)) in selected_stickies
                if len(selected_stickies) > 1 and is_obj_selected:
                    menu.add_command(
                        label="まとめてPDFを作成",
                        command=self._convert_selected_stickies_pipeline,
                    )

            elif (
                type_ == "connection"
                and obj["from_type"] == "note"
                and obj["to_type"] == "note"
            ):
                is_linked = self._check_db_link_exists(obj["from_key"], obj["to_key"])
                if is_linked:
                    menu.add_command(
                        label="リンク解除",
                        command=lambda: self._handle_connection_double_click(obj),
                    )
                else:
                    menu.add_command(
                        label="リンク作成",
                        command=lambda: self._handle_connection_double_click(obj),
                    )

            menu.add_command(
                label="削除", command=lambda: self._delete_item(type_, obj)
            )
            menu.post(event.x_root, event.y_root)

    def _handle_connection_double_click(self, conn_data):
        key_a = conn_data["from_key"]
        key_b = conn_data["to_key"]
        title_a = self.notes_on_canvas[key_a]["title"]
        title_b = self.notes_on_canvas[key_b]["title"]

        is_linked = self._check_db_link_exists(key_a, key_b)

        if is_linked:
            msg = f"以下のノート間のリンクを解除しますか？\n\n・{title_a}\n・{title_b}"
            if messagebox.askyesno("リンク解除", msg, parent=self):
                self.focus_force()
                self._perform_db_link_removal(key_a, key_b, conn_data["id"])
            else:
                self.focus_force()
        else:
            msg = f"以下のノート間に相互リンクを作成しますか？\n\n・{title_a}\n・{title_b}"
            if messagebox.askyesno("リンク作成", msg, parent=self):
                self.focus_force()
                self._perform_db_link_update(key_a, key_b, conn_data["id"])
            else:
                self.focus_force()

    def _perform_db_link_update(self, key_a, key_b, item_id):
        if not self.parent_app.loaded_db_path:
            return
        conn = None
        try:
            conn = sqlite3.connect(self.parent_app.loaded_db_path)
            cursor = conn.cursor()
            self._append_link_text(
                cursor, key_a, key_b, self.notes_on_canvas[key_b]["title"]
            )
            self._append_link_text(
                cursor, key_b, key_a, self.notes_on_canvas[key_a]["title"]
            )
            conn.commit()
            self.canvas.itemconfigure(item_id, fill="#28a745")
            messagebox.showinfo("完了", "リンクを作成しました。", parent=self)
        except Exception as e:
            if conn:
                conn.rollback()
            messagebox.showerror("DBエラー", f"失敗: {e}", parent=self)
        finally:
            if conn:
                conn.close()
            self.focus_force()

    def _perform_db_link_removal(self, key_a, key_b, item_id):
        if not self.parent_app.loaded_db_path:
            return
        conn = None
        try:
            conn = sqlite3.connect(self.parent_app.loaded_db_path)
            cursor = conn.cursor()
            self._remove_link_text_from_db(cursor, key_a, key_b)
            self._remove_link_text_from_db(cursor, key_b, key_a)
            conn.commit()
            base_color = "white" if self.bg_color == "#2b2b2b" else "black"
            self.canvas.itemconfigure(item_id, fill=base_color)
            messagebox.showinfo("完了", "リンクを解除しました。", parent=self)
        except Exception as e:
            if conn:
                conn.rollback()
            messagebox.showerror("DBエラー", f"失敗: {e}", parent=self)
        finally:
            if conn:
                conn.close()
            self.focus_force()

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

    # --- Export Pipeline ---
    def _convert_sticky_to_note_pipeline(self, sticky_obj):
        default_title = sticky_obj.get("title", "NOTITLE")

        if hasattr(self.parent_app, "commonplace_keys_options"):
            key_options = self.parent_app.commonplace_keys_options
        else:
            key_options = []

        dialog = ConversionDialog(self, default_title, key_options)
        self.wait_window(dialog)

        if dialog.result:
            title, key = dialog.result
            content = sticky_obj.get("content", "")
            bg_color = sticky_obj.get("bg_color", "#FFFFA5")
            self._process_md_pdf_creation(title, content, key, bg_color)

    def _convert_selected_stickies_pipeline(self):
        combined_content = ""
        targets = []
        for t, k in self.selected_items:
            if t == "sticky":
                s = next((x for x in self.stickies_on_canvas if str(id(x)) == k), None)
                if s:
                    targets.append(s)

        if not targets:
            return

        if hasattr(self.parent_app, "commonplace_keys_options"):
            key_options = self.parent_app.commonplace_keys_options
        else:
            key_options = []

        dialog = ConversionDialog(self, "まとめノート", key_options)
        self.wait_window(dialog)

        if dialog.result:
            title, key = dialog.result
            for s in targets:
                t_ = s.get("title", "NOTITLE")
                c_ = s.get("content", "")
                combined_content += f"## {t_}\n{c_}\n\n"

            self._process_md_pdf_creation(title, combined_content, key, "#FFFFA5")

    def _process_md_pdf_creation(self, title, content, index_key, bg_color):
        now_str = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
        safe_title = re.sub(r'[\\/:\*\?"<>\|]', "_", title if title else "NOTITLE")
        base_name = f"{now_str}_{safe_title}"

        # 設定した nexus_output_folder を使用
        if hasattr(self.parent_app, "nexus_output_folder"):
            save_dir = self.parent_app.nexus_output_folder
        else:
            save_dir = Path("Nexus_Output")

        if not save_dir.exists():
            try:
                save_dir.mkdir(parents=True, exist_ok=True)
            except Exception as e:
                messagebox.showerror(
                    "エラー",
                    f"保存先フォルダを作成できません:\n{save_dir}\n{e}",
                    parent=self,
                )
                return

        pdf_path = save_dir / f"{base_name}.pdf"

        try:
            # ▼ 変更箇所: 一時フォルダ作成を前倒しし、mdファイルをそこに保存 ▼
            temp_dir = save_dir / "temp_canvas_process"
            temp_dir.mkdir(exist_ok=True)

            # MDファイルを一時フォルダ内に作成 (処理後に削除される)
            md_path = temp_dir / f"{base_name}.md"

            style_tag = (
                "<style> body "
                f"{{ background-color: {bg_color}; padding: 20px; }} "
                "</style>"
            )
            meta_comment = "\n\n"

            md_text = (
                f"{style_tag}\n{meta_comment}\n" f"# {title}\n\n" f"# [内容]\n{content}"
            )

            with open(md_path, "w", encoding="utf-8") as f:
                f.write(md_text)

            # temp_pdf, temp_flat のパス定義
            temp_pdf = temp_dir / f"temp_{base_name}.pdf"
            temp_flat = temp_dir / f"flat_{base_name}.pdf"

            # self.font_path を使用 (config_dataエラー回避)
            if self.font_path:
                font_path = self.font_path
            else:
                font_path = r"C:\Windows\Fonts\msgothic.ttc"

            convert_document_to_pdf(md_path, temp_pdf, paper_size_str="A4")
            high_fidelity_flatten(
                str(temp_pdf), str(temp_flat), str(font_path), flatten_ink=False
            )
            normalize_pdf_to_papersize(
                str(temp_flat), str(pdf_path), 595.276, 841.89, target_format="A4"
            )
            embed_processing_flag(str(pdf_path))

            # QRコードの埋め込み
            key_rect_str = self._get_config_value(
                "Extraction", "key_rect", "0, 13, 391, 73"
            )
            try:
                key_rect_tuple = tuple(map(float, key_rect_str.split(",")))
            except Exception:
                key_rect_tuple = (0, 13, 391, 73)

            add_metadata_to_clip(
                pdf_path_str=str(pdf_path),
                font_path=str(font_path),
                paper_width=595.276,  # A4
                paper_height=841.89,
                key_rect_tuple=key_rect_tuple,
                index_key_to_embed=index_key,
                text_color=(0, 0, 0),
                comment_to_embed="",
                base_name=base_name,  # ユニークID生成用
            )

            shutil.rmtree(temp_dir)  # ここでMDファイルごと一時フォルダを削除

            messagebox.showinfo(
                "完了", f"ファイルを生成しました:\n{pdf_path.name}", parent=self
            )

        except Exception as e:
            messagebox.showerror("エラー", f"処理に失敗しました:\n{e}", parent=self)

    def _delete_item(self, type_, obj):
        if type_ == "sticky":
            self.canvas.delete(obj["ids"][0])
            self.canvas.delete(obj["ids"][1])
            self.canvas.delete(obj["shadow_id"])

            self.stickies_on_canvas.remove(obj)
            sid = str(id(obj))
            to_rem = [
                c
                for c in self.connections_on_canvas
                if (
                    (c["from_type"] == "sticky" and c["from_key"] == sid)
                    or (c["to_type"] == "sticky" and c["to_key"] == sid)
                )
            ]
            for c in to_rem:
                self.canvas.delete(c["id"])
                self.connections_on_canvas.remove(c)
        elif type_ == "note":
            self._delete_note(obj)
        elif type_ == "connection":
            self.canvas.delete(obj["id"])
            self.connections_on_canvas.remove(obj)
        elif type_ == "shape":
            self.canvas.delete(obj["id"])
            self.shapes_on_canvas.remove(obj)
        self.save_canvas()

    def _delete_note(self, key):
        ids = self.notes_on_canvas[key]["ids"]
        for i in ids:
            self.canvas.delete(i)
        del self.notes_on_canvas[key]
        if ("note", key) in self.selected_items:
            self.selected_items.remove(("note", key))
        to_rem = [
            c
            for c in self.connections_on_canvas
            if (c["from_key"] == key or c["to_key"] == key)
        ]
        for c in to_rem:
            self.canvas.delete(c["id"])
            self.connections_on_canvas.remove(c)

    # --- ファイルIO ---
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

    def _save_to_json(self, path):
        data = {
            "notes": {
                k: {
                    "x": v["x"],
                    "y": v["y"],
                    "title": v["title"],
                    "cp_key": v["cp_key"],
                }
                for k, v in self.notes_on_canvas.items()
            },
            "stickies": [
                {
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
            del item["id"]
            data["shapes"].append(item)

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
        if not path.exists():
            return
        self.clear_canvas_items()
        try:
            with open(path, "r", encoding="utf-8") as f:
                data = json.load(f)
            for k, v in data.get("notes", {}).items():
                self.create_note_item(
                    k, v["title"], v.get("cp_key", ""), v["x"], v["y"]
                )

            for s in data.get("stickies", []):
                t = s.get("title", "")
                c = s.get("content", "")
                if not t and not c and "text" in s:
                    old_text = s["text"]
                    match = re.match(r"^(\S+)\s*(.*)$", old_text, re.DOTALL)
                    if match:
                        t, c = match.group(1), match.group(2)
                    else:
                        t = "NOTITLE"
                        c = old_text
                # JSONに保存されている値はスケール1.0基準なので、現在のスケールを掛ける
                loaded_w = s.get("w", 180) * self.current_scale
                loaded_h = s.get("h", 120) * self.current_scale

                self.create_sticky_item(
                    s["x"],
                    s["y"],
                    t,
                    c,
                    bg_color=s.get("bg_color", "#FFFFA5"),
                    w=loaded_w,
                    h=loaded_h,
                )

            for s in data.get("shapes", []):
                if s["type"] == "line":
                    self.create_shape_item(
                        "line", x=s["x1"], y=s["y1"], x2=s["x2"], y2=s["y2"]
                    )
                elif s["type"] == "rect":
                    self.create_shape_item(
                        s["type"], s["x"], s["y"], s.get("w", 0), s.get("h", 0)
                    )

            for c in data.get("connections", []):
                ft = c.get("from_type", "note")
                tt = c.get("to_type", "note")
                self.create_connection_item((ft, c["from_key"]), (tt, c["to_key"]))
            self.canvas.configure(scrollregion=self.canvas.bbox("all"))
        except Exception as e:
            logger.error(f"Canvas load error: {e}")
            messagebox.showerror("エラー", f"読込失敗: {e}", parent=self)

    def clear_canvas(self):
        if messagebox.askyesno("確認", "キャンバスをクリアしますか？", parent=self):
            self.clear_canvas_items()
            self.save_canvas()

    def clear_canvas_items(self):
        self.canvas.delete("all")
        self._draw_grid(width=5000, height=5000)
        self.notes_on_canvas = {}
        self.stickies_on_canvas = []
        self.shapes_on_canvas = []
        self.connections_on_canvas = []
        self.selected_items = set()
        self.current_scale = 1.0
        if hasattr(self, "zoom_label_var"):
            self.zoom_label_var.set("100%")
        self.center_view()

    def export_canvas_dialog(self):
        file_path = filedialog.asksaveasfilename(
            defaultextension=".pdf",
            filetypes=[("PDF Document", "*.pdf")],
            title="PDFとして保存",
            parent=self,
        )
        if not file_path:
            return
        try:
            self._export_to_file(Path(file_path))
            messagebox.showinfo(
                "完了", f"出力が完了しました:\n{Path(file_path).name}", parent=self
            )
        except Exception as e:
            messagebox.showerror("エラー", f"出力に失敗しました:\n{e}", parent=self)
            logger.error(f"PDF Export Error: {e}")

    def _hex_to_rgb(self, hex_color):
        hex_color = hex_color.lstrip("#")
        return tuple(int(hex_color[i : i + 2], 16) / 255.0 for i in (0, 2, 4))

    def _export_to_file(self, output_path):
        """
        キャンバスの内容を一時PDFに出力し、Normalisiererの機能を使って
        指定用紙サイズ（A4等）に正規化して保存する。
        """
        # --- 1. コンテンツ領域の計算 ---
        min_x, min_y, max_x, max_y = (
            float("inf"),
            float("inf"),
            float("-inf"),
            float("-inf"),
        )

        def update_bounds(x, y, w, h):
            nonlocal min_x, min_y, max_x, max_y
            min_x = min(min_x, x)
            min_y = min(min_y, y)
            max_x = max(max_x, x + w)
            max_y = max(max_y, y + h)

        # バウンディングボックス計算
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

        # コンテンツサイズ + マージン
        margin = 50
        content_w = max_x - min_x + margin * 2
        content_h = max_y - min_y + margin * 2

        # --- 2. 一時PDFの生成 (Normalisiererに渡すための原版) ---
        # 一時ディレクトリを作成して処理
        with tempfile.TemporaryDirectory() as temp_dir:
            temp_raw_pdf = Path(temp_dir) / "raw_canvas_export.pdf"

            doc = fitz.open()
            page = doc.new_page(width=content_w, height=content_h)

            # メタデータ設定
            doc.set_metadata(
                {
                    "keywords": "Synapsen:Whiteboard",
                    "creator": "Synapsen Canvas",
                }
            )

            try:
                page.insert_font(
                    fontname="msgothic", fontfile=r"C:\Windows\Fonts\msgothic.ttc"
                )
            except Exception:
                pass

            shape = page.new_shape()

            # 座標変換用 (左上原点 + マージン)
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
                    s = next(
                        (x for x in self.stickies_on_canvas if str(id(x)) == key_), None
                    )
                    if s:
                        return tx(s["x"]) + s["w"] / 2, ty(s["y"]) + s["h"] / 2
                return None, None

            # --- 描画 (回転なし・単純描画) ---

            # 1. 接続線
            for c in self.connections_on_canvas:
                x1, y1 = get_center(c["from_type"], c["from_key"])
                x2, y2 = get_center(c["to_type"], c["to_key"])

                if x1 is not None and x2 is not None:
                    is_linked = False
                    if c["from_type"] == "note" and c["to_type"] == "note":
                        is_linked = self._check_db_link_exists(
                            c["from_key"], c["to_key"]
                        )
                    color = (0.15, 0.65, 0.27) if is_linked else (0, 0, 0)

                    shape.draw_line(fitz.Point(x1, y1), fitz.Point(x2, y2))
                    shape.finish(color=color, width=1)

            # 2. 図形
            for s in self.shapes_on_canvas:
                if s["type"] == "rect":
                    rtx, rty = tx(s["x"]), ty(s["y"])
                    shape.draw_rect(
                        fitz.Rect(rtx, rty, rtx + s.get("w", 0), rty + s.get("h", 0))
                    )
                    shape.finish(color=(1, 0, 0), width=1, dashes=[4, 4])
                elif s["type"] == "line":
                    x1, y1 = tx(s["x1"]), ty(s["y1"])
                    x2, y2 = tx(s["x2"]), ty(s["y2"])
                    shape.draw_line(fitz.Point(x1, y1), fitz.Point(x2, y2))
                    shape.finish(color=(0, 0, 0), width=1)

            # 3. 付箋
            for s in self.stickies_on_canvas:
                rtx, rty = tx(s["x"]), ty(s["y"])
                w, h = s["w"], s["h"]

                # 影
                shape.draw_rect(fitz.Rect(rtx + 5, rty + 5, rtx + w + 5, rty + h + 5))
                shape.finish(fill=(0.5, 0.5, 0.5), stroke_opacity=0)

                # 本体
                shape.draw_rect(fitz.Rect(rtx, rty, rtx + w, rty + h))
                shape.finish(
                    fill=self._hex_to_rgb(s["bg_color"]), color=(0, 0, 0), width=0
                )

                # テキスト
                disp = self._create_sticky_display_text(
                    s.get("title", ""), s.get("content", "")
                )
                shape.insert_textbox(
                    fitz.Rect(rtx + 5, rty + 5, rtx + w - 5, rty + h - 5),
                    disp,
                    fontname="msgothic",
                    fontsize=12,
                    color=(0, 0, 0),
                )

            # 4. ノート
            for key, info in self.notes_on_canvas.items():
                rtx, rty = tx(info["x"]), ty(info["y"])
                col = self.parent_app.key_colors.get(info["cp_key"].lower(), "#aaaaaa")

                shape.draw_rect(fitz.Rect(rtx, rty, rtx + 160, rty + 80))
                shape.finish(fill=self._hex_to_rgb(col), color=(0, 0, 0), width=1)

                shape.insert_textbox(
                    fitz.Rect(rtx + 5, rty + 5, rtx + 155, rty + 75),
                    f"[[{key}: {info['title']}]]",
                    fontname="msgothic",
                    fontsize=10,
                    align=1,
                    color=(0, 0, 0),
                )

            shape.commit()
            doc.save(str(temp_raw_pdf))
            doc.close()

            # --- 3. Normalisiererによる正規化処理 (A4化) ---
            # 親アプリの設定から用紙サイズを取得
            target_format = self.parent_app.config_data.get("paper_size", "A4")
            paper_width = self.parent_app.paper_width
            paper_height = self.parent_app.paper_height

            # pdf_utilsのnormalize_pdf_to_papersizeを使用
            # これにより、コンテンツ全体が指定用紙サイズ(A4/A5)の中央にフィットするようにリサイズされます
            normalize_pdf_to_papersize(
                str(temp_raw_pdf),
                str(output_path),
                paper_width,
                paper_height,
                target_format=target_format,
            )

            # 処理済みフラグを付与 (Normalisiererでの再処理防止)
            embed_processing_flag(str(output_path))

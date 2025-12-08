import re
import io
import os
import customtkinter as ctk
from tkinter import filedialog, messagebox
from pathlib import Path
import datetime
import tempfile
import shutil
import fitz
from urllib.parse import urlparse, unquote
import urllib.request
from PIL import Image, ImageTk

from pdf_utils import add_metadata_to_clip, hex_to_rgb_tuple

import logging

logger = logging.getLogger(__name__)

# --- Playwright パス設定 (EXE化対策) ---
# PyInstallerでバンドルされた環境でも、システムにインストールされたブラウザを見に行くように強制する
if "LOCALAPPDATA" in os.environ:
    # Windowsの標準パス: C:\Users\User\AppData\Local\ms-playwright
    playwright_browsers_path = Path(os.environ["LOCALAPPDATA"]) / "ms-playwright"
    os.environ["PLAYWRIGHT_BROWSERS_PATH"] = str(playwright_browsers_path)
else:
    # フォールバック (macOS/Linux等は 0 で標準パスを見る)
    os.environ["PLAYWRIGHT_BROWSERS_PATH"] = "0"

sync_playwright = None
PlaywrightError = Exception
PlaywrightTimeoutError = Exception

try:
    from playwright.sync_api import sync_playwright, Error, TimeoutError

    PlaywrightError = Error
    PlaywrightTimeoutError = TimeoutError
except ImportError:
    pass


class PreviewWindow(ctk.CTkToplevel):
    """
    WebClipのプレビューを表示するための専用ウィンドウ。
    Canvasを使用し、縦横スクロールに対応。
    """

    def __init__(self, parent, pil_image, icon_path=None):
        super().__init__(parent)
        self.title("ページプレビュー")
        self.geometry("800x600")
        self._custom_icon_path = icon_path

        self.transient(parent)
        self.grab_set()

        # グリッド設定 (キャンバスを最大化)
        self.grid_rowconfigure(0, weight=1)
        self.grid_columnconfigure(0, weight=1)

        # --- キャンバスとスクロールバーの配置 ---
        # 背景色はダークグレーにしておく (画像が見やすいように)
        self.canvas = ctk.CTkCanvas(self, bg="#2b2b2b", highlightthickness=0)
        self.canvas.grid(row=0, column=0, sticky="nsew")

        # 縦スクロールバー
        self.vsb = ctk.CTkScrollbar(
            self, orientation="vertical", command=self.canvas.yview
        )
        self.vsb.grid(row=0, column=1, sticky="ns")

        # 横スクロールバー
        self.hsb = ctk.CTkScrollbar(
            self, orientation="horizontal", command=self.canvas.xview
        )
        self.hsb.grid(row=1, column=0, sticky="ew")

        # スクロールコマンドの紐付け
        self.canvas.configure(yscrollcommand=self.vsb.set, xscrollcommand=self.hsb.set)

        # --- 画像の描画 ---
        # ガベージコレクション対策でインスタンス変数に保持
        self.tk_image = ImageTk.PhotoImage(pil_image)

        # キャンバスに画像を描画 (左上基準)
        self.canvas.create_image(0, 0, anchor="nw", image=self.tk_image)

        # スクロール範囲を画像サイズに合わせる
        self.canvas.configure(scrollregion=(0, 0, pil_image.width, pil_image.height))

        # --- マウスホイール操作のバインド ---
        self.canvas.bind("<MouseWheel>", self._on_mousewheel)
        self.canvas.bind("<Shift-MouseWheel>", self._on_shift_mousewheel)

        # アイコン設定
        if self._custom_icon_path:
            self.after(200, lambda: self.iconbitmap(default=self._custom_icon_path))

        self.lift()
        self.focus_force()

    def _on_mousewheel(self, event):
        """縦スクロール (ホイール)"""
        self.canvas.yview_scroll(int(-1 * (event.delta / 120)), "units")

    def _on_shift_mousewheel(self, event):
        """横スクロール (Shift + ホイール)"""
        self.canvas.xview_scroll(int(-1 * (event.delta / 120)), "units")

    def iconbitmap(self, *args, **kwargs):
        """アイコン設定の強制オーバーライド"""
        if self._custom_icon_path:
            try:
                super().iconbitmap(self._custom_icon_path)
            except Exception:
                pass
        else:
            try:
                super().iconbitmap(*args, **kwargs)
            except Exception:
                pass


class WebClipWindow(ctk.CTkToplevel):
    """
    URLを指定してWebページをPDFとしてクリップし、
    IndexKey, コメント, 書誌情報を埋め込むためのToplevelウィンドウ（モーダル）。

    Playwrightライブラリを使用してWebページの取得とPDF化を行います。

    Attributes:
        parent_app (ctk.CTk): 親である Synapsen_Normalisierer のインスタンス。
        temp_dir (Path | None): 一時PDFの保存先ディレクトリ。
        playwright_context: Playwright の sync_playwright コンテキスト。
        browser: Playwright のブラウザインスタンス。
        page: Playwright のページインスタンス。
    """

    def __init__(self, parent_app):
        """
        WebClipWindowを初期化します。

        Args:
            parent_app (ctk.CTk): 親ウィンドウ (Synapsen_Normalisierer)。
        """
        super().__init__(parent_app)

        self.parent_app = parent_app
        self.temp_dir = None
        self.page_title_cache = ""
        self.site_name_cache = ""
        self.fetched_content_type = None

        # --- Playwright (実行コンテキスト) ---
        self.playwright_context = None
        self.browser = None
        self.page = None

        # --- Playwrightがインストールされているかチェック ---
        if sync_playwright is None:
            messagebox.showerror(
                "ライブラリ不足エラー",
                "Webクリップ機能に必要な 'playwright' ライブラリが見つかりません。\n\n"
                "この機能を使用する場合は、`install_option.bat` 又は `install.bat`を使用して、"
                " 'playwright' ライブラリをインストールして下さい\n"
                "詳細は READMEの「実行方法」の項目をご参照ください。",
                parent=parent_app,
            )
            # ウィンドウの表示シーケンスから抜け、即座に閉じる
            self.after(100, self.destroy)
            return

        self.title("Webクリップで正規化")
        self.geometry("450x830")

        # アイコン設定
        self._custom_icon_path = None
        if hasattr(parent_app, "icon_path") and parent_app.icon_path:
            self._custom_icon_path = str(parent_app.icon_path)
            self.after(200, lambda: self.iconbitmap(default=self._custom_icon_path))

        self.protocol("WM_DELETE_WINDOW", self.on_close)
        self.transient(parent_app)
        self.grab_set()

        # --- [UI定義] ---

        # --- 1. URLとファイル名 ---
        input_frame = ctk.CTkFrame(self, fg_color="gray25")
        input_frame.pack(pady=10, padx=10, fill="x")

        ctk.CTkLabel(input_frame, text="URL:", width=80).grid(
            row=0, column=0, padx=5, pady=5, sticky="w"
        )
        self.url_entry = ctk.CTkEntry(input_frame, placeholder_text="https://...")
        self.url_entry.grid(row=0, column=1, columnspan=2, padx=5, pady=5, sticky="ew")

        ctk.CTkLabel(input_frame, text="ファイル名:", width=80).grid(
            row=1, column=0, padx=5, pady=5, sticky="w"
        )
        now = datetime.datetime.now()
        default_base_name = f"{now.strftime('%Y%m%d_%H%M%S')}_WebClip"
        self.filename_var = ctk.StringVar(value=default_base_name)
        self.filename_entry = ctk.CTkEntry(input_frame, textvariable=self.filename_var)
        self.filename_entry.grid(row=1, column=1, padx=5, pady=5, sticky="ew")
        ctk.CTkLabel(input_frame, text=".pdf").grid(
            row=1, column=2, padx=5, pady=5, sticky="w"
        )

        input_frame.grid_columnconfigure(1, weight=1)

        # --- 2. ページ情報取得ボタン ---
        self.fetch_button = ctk.CTkButton(
            self, text="1. ページ情報取得 (書誌情報)", command=self.fetch_page_info
        )
        self.fetch_button.pack(pady=5, padx=10, fill="x")

        # プレビューボタン (前回の追加分)
        self.preview_button = ctk.CTkButton(
            self,
            text="内容確認 (スクリーンショット)",
            command=self.show_page_preview,
            state="disabled",
            fg_color="#585a9c",
        )
        self.preview_button.pack(pady=2, padx=10, fill="x")

        # --- 3. 書誌情報 (SIST 02) 編集フレーム ---
        ctk.CTkLabel(self, text="書誌情報 (SIST 02準拠)", anchor="w").pack(
            pady=(10, 0), padx=10, fill="x"
        )
        sist_frame = ctk.CTkFrame(self)
        sist_frame.pack(pady=(0, 10), padx=10, fill="x")
        sist_frame.grid_columnconfigure(1, weight=1)

        ctk.CTkLabel(sist_frame, text="著者名:").grid(
            row=0, column=0, padx=5, pady=5, sticky="w"
        )
        self.sist_author_entry = ctk.CTkEntry(
            sist_frame, placeholder_text="（自動取得試行）"
        )
        self.sist_author_entry.grid(row=0, column=1, padx=5, pady=5, sticky="ew")

        ctk.CTkLabel(sist_frame, text="ページ名:").grid(
            row=1, column=0, padx=5, pady=5, sticky="w"
        )
        self.sist_title_entry = ctk.CTkEntry(
            sist_frame, placeholder_text="（自動取得試行）"
        )
        self.sist_title_entry.grid(row=1, column=1, padx=5, pady=5, sticky="ew")

        ctk.CTkLabel(sist_frame, text="サイト名:").grid(
            row=2, column=0, padx=5, pady=5, sticky="w"
        )
        self.sist_site_entry = ctk.CTkEntry(
            sist_frame, placeholder_text="（自動取得試行）"
        )
        self.sist_site_entry.grid(row=2, column=1, padx=5, pady=5, sticky="ew")

        ctk.CTkLabel(sist_frame, text="更新日:").grid(
            row=3, column=0, padx=5, pady=5, sticky="w"
        )
        self.sist_date_entry = ctk.CTkEntry(
            sist_frame, placeholder_text="（自動取得試行, YYYY-MM-DD）"
        )
        self.sist_date_entry.grid(row=3, column=1, padx=5, pady=5, sticky="ew")

        # --- 4. IndexKey 選択 ---
        ctk.CTkLabel(self, text="IndexKey (PDF 1ページ目に埋込)", anchor="w").pack(
            pady=(10, 0), padx=10, fill="x"
        )
        key_frame = ctk.CTkFrame(self)
        key_frame.pack(pady=(0, 10), padx=10, fill="x")

        key_options = self.parent_app.config_data.get("commonplace_keys_options", [])

        self.index_key_combo = ctk.CTkComboBox(
            key_frame, values=["（未選択）"] + key_options
        )
        self.index_key_combo.set("（未選択）")
        self.index_key_combo.pack(fill="x", padx=5, pady=5)

        # --- 5. コメント入力 ---
        ctk.CTkLabel(self, text="コメント (PDF 最終ページに埋込)", anchor="w").pack(
            pady=(10, 0), padx=10, fill="x"
        )
        comment_frame = ctk.CTkFrame(self)
        comment_frame.pack(pady=(0, 10), padx=10, fill="both", expand=True)
        self.comment_textbox = ctk.CTkTextbox(comment_frame, height=80)
        self.comment_textbox.pack(fill="both", expand=True, padx=5, pady=5)

        # --- 7. Key入力欄 ---
        ctk.CTkLabel(
            self, text="引用元Key (カンマ区切り または 改行区切り)", anchor="w"
        ).pack(pady=(10, 0), padx=10, fill="x")
        key_frame2 = ctk.CTkFrame(self)

        self.cited_keys_entry = ctk.CTkTextbox(key_frame2, height=60)
        self.cited_keys_entry.pack(fill="both", expand=True, padx=5, pady=5)
        key_frame2.pack(pady=(0, 10), padx=10, fill="x")

        # --- 8. 実行ボタン ---
        self.run_button = ctk.CTkButton(
            self,
            text="2. 出力先を選んでクリップ実行",
            command=self.run_webclip_process,
            state="disabled",  # 初期状態は無効
        )
        self.run_button.pack(pady=10, padx=10, fill="x", ipady=10)

        # --- 9. ステータスラベル ---
        self.status_label = ctk.CTkLabel(self, text="")
        self.status_label.pack(pady=10, padx=10)

    def on_close(self) -> None:
        """
        ウィンドウが閉じられるとき（[x]ボタン押下時）の処理。
        Playwrightのセッションを安全に終了し、一時フォルダをクリーンアップします。
        """
        # Playwrightセッションの終了
        if self.page:
            try:
                self.page.close()
            except Exception as e:
                logger.error(f"Playwright page close error: {e}")
        if self.browser:
            try:
                self.browser.close()
            except Exception as e:
                logger.error(f"Playwright browser close error: {e}")
        if self.playwright_context:
            try:
                self.playwright_context.stop()
            except Exception as e:
                logger.error(f"Playwright context stop error: {e}")

        # 一時フォルダのクリーンアップ
        if self.temp_dir and self.temp_dir.exists():
            try:
                shutil.rmtree(self.temp_dir)
            except Exception as e:
                logger.error(f"WebClip一時フォルダの削除に失敗: {e}")

        self.grab_release()
        self.destroy()

    def iconbitmap(self, *args, **kwargs):
        """
        CustomTkinterがアイコンをリセットするのを防ぐためのオーバーライドメソッド。
        """
        if self._custom_icon_path:
            try:
                super().iconbitmap(self._custom_icon_path)
            except Exception:
                pass
        else:
            try:
                super().iconbitmap(*args, **kwargs)
            except Exception:
                pass

    def fetch_page_info(self) -> None:
        """
        「1. ページ情報取得」ボタンの処理。
        URLのContent-Typeを判別し、HTMLの場合はPlaywrightで書誌情報を、
        それ以外の場合やタイムアウトやエラーが発生した場合は、限定的な情報を取得します。
        """
        url = self.url_entry.get().strip()
        if not url.startswith("http://") and not url.startswith("https://"):
            messagebox.showerror(
                "入力エラー",
                "有効なURL (https://...) を入力してください。",
                parent=self,
            )
            return

        self.status_label.configure(text="ページ情報を取得中...", text_color="gray")
        self.fetch_button.configure(state="disabled")
        self.run_button.configure(state="disabled")
        self.update_idletasks()

        self.fetched_content_type = "text/html"
        try:
            # ユーザーエージェントを偽装して403エラーを回避
            headers = {"User-Agent": "Mozilla/5.0"}
            req = urllib.request.Request(url, headers=headers)
            with urllib.request.urlopen(req, timeout=10) as response:
                self.fetched_content_type = response.getheader(
                    "Content-Type", "text/html"
                ).lower()
        except Exception:
            pass

        try:
            if "text/html" not in self.fetched_content_type:
                # PDF/画像等の場合
                parsed_url = urlparse(url)
                # URLデコード (例: %E3%81... -> 日本語)
                filename_from_url = unquote(Path(parsed_url.path).name)

                self.sist_title_entry.delete(0, "end")
                self.sist_title_entry.insert(
                    0, filename_from_url or "ダウンロードファイル"
                )
                self.sist_site_entry.delete(0, "end")
                self.sist_site_entry.insert(0, parsed_url.netloc)
                self.status_label.configure(
                    text="PDF/画像リンクを検出。", text_color="gray"
                )
                self.fetch_button.configure(state="normal")
                self.run_button.configure(state="normal")
                return

            # HTMLの場合
            self.status_label.configure(
                text="ページ情報を取得中 (最大1分)...", text_color="gray"
            )
            self.update_idletasks()

            # 1. Playwrightセッションがなければ開始する
            if self.playwright_context is None:
                self.playwright_context = sync_playwright().start()
                self.browser = self.playwright_context.chromium.launch()
                self.page = self.browser.new_page()

            # 2. ページに移動 (タイムアウト1分)
            self.page.goto(url, wait_until="domcontentloaded", timeout=60000)
            self.preview_button.configure(state="normal")  # Previewの有効化

            self.page_title_cache = self.page.title()

            # メタデータ抽出
            author = self.page.locator(
                'meta[name="author"], meta[property="og:author"]'
            ).first.get_attribute("content")

            site_name = self.page.locator(
                'meta[property="og:site_name"]'
            ).first.get_attribute("content")

            if not site_name:
                site_name = urlparse(url).netloc
            self.site_name_cache = site_name

            date_str = self.page.locator(
                'meta[property="article:published_time"], time[datetime]'
            ).first.get_attribute("datetime")

            # 'YYYY-MM-DD' の部分のみ取得
            iso_date = date_str.split("T")[0] if date_str else ""

            # 4. UIのEntryに取得した情報をセット
            self.sist_title_entry.delete(0, "end")
            self.sist_title_entry.insert(0, self.page_title_cache)
            self.sist_site_entry.delete(0, "end")
            self.sist_site_entry.insert(0, self.site_name_cache)

            if author:
                self.sist_author_entry.delete(0, "end")
                self.sist_author_entry.insert(0, author)
            if iso_date:
                self.sist_date_entry.delete(0, "end")
                self.sist_date_entry.insert(0, iso_date)

            self.status_label.configure(text="情報を取得しました。")

        except PlaywrightTimeoutError as e:
            # --- タイムアウトの場合 (フォールバック準備) ---
            logger.warning(f"ページ読み込みタイムアウト: {e}")
            self.status_label.configure(
                text="タイムアウト。簡易情報(タイトル/ドメイン)を取得します...",
                text_color="orange",
            )

            try:
                # タイムアウトしてもタイトル等が取れる場合がある
                self.page_title_cache = self.page.title()
                parsed_url = urlparse(url)
                self.site_name_cache = parsed_url.netloc

                self.sist_title_entry.delete(0, "end")
                self.sist_title_entry.insert(0, self.page_title_cache)
                self.sist_site_entry.delete(0, "end")
                self.sist_site_entry.insert(0, self.site_name_cache)

                messagebox.showwarning(
                    "情報取得タイムアウト",
                    f"ページの読み込みがタイムアウトしました (1分)。\n{e}\n\n"
                    "タイトルとサイト名のみ取得を試みました。\n"
                    "著者名や更新日は手動で入力してください。",
                    parent=self,
                )
            except Exception:
                messagebox.showerror(
                    "情報取得エラー",
                    "タイムアウト後の簡易情報取得にも失敗しました:\n",
                    parent=self,
                )
                self.status_label.configure(
                    text="情報取得失敗（タイムアウト）。手動で入力してください。",
                    text_color="orange",
                )
                self.page_title_cache = ""
                self.site_name_cache = ""

        except Exception as e:
            # --- その他のエラー (ブラウザ無し等) ---
            logger.error(f"ページ読み込みエラー: {e}")
            messagebox.showerror(
                "情報取得エラー",
                f"エラーが発生しました:\n{e}\n\nURLやブラウザのインストール状況を確認してください。",
                parent=self,
            )
            # オブジェクト破棄
            if self.page:
                try:
                    self.page.close()
                except Exception:
                    pass
                self.page = None
            if self.browser:
                try:
                    self.browser.close()
                except Exception:
                    pass
                self.browser = None

        finally:
            self.fetch_button.configure(state="normal")
            self.run_button.configure(state="normal")  # 実行ボタンを有効化

    def show_page_preview(self):
        """現在のPlaywrightページのスクリーンショットを表示する"""
        if not self.page:
            return

        try:
            # スクリーンショットをメモリ上に取得
            png_bytes = self.page.screenshot(full_page=True)
            pil_image = Image.open(io.BytesIO(png_bytes))

            # インスタンス生成時に pil_image を渡す
            PreviewWindow(self, pil_image, icon_path=self._custom_icon_path)

        except Exception as e:
            messagebox.showerror(
                "プレビューエラー", f"プレビューの生成に失敗しました:\n{e}", parent=self
            )

    def run_webclip_process(self) -> None:
        """
        「2. クリップ実行」ボタンの処理。

        Content-Typeに応じてHTMLはPlaywrightでPDF化、
        その他(PDF/画像)は直接ダウンロードし、
        その後、共通の正規化処理とメタデータ埋め込みを行います。
        """
        url = self.url_entry.get().strip()

        # --- 1. 入力バリデーション ---
        # base_name の取得とサニタイズ
        base_name_raw = self.filename_var.get().strip()
        base_name = re.sub(r'[\\/:\*\?"<>\|]', "_", base_name_raw)

        # もし置換が発生したら、UI (StringVar) にも反映する
        if base_name != base_name_raw:
            self.filename_var.set(base_name)

        if not url:
            messagebox.showerror(
                "入力エラー",
                "有効なURL (https://...) を入力してください。",
                parent=self,
            )
            return
        if not base_name:
            messagebox.showerror(
                "入力エラー", "ファイル名を入力してください。", parent=self
            )
            return

        font_path = self.parent_app.font_path
        if not font_path or not Path(font_path).is_file():
            self.parent_app.status_label.configure(
                text="configエラー: フォント設定", text_color="orange"
            )
            return

        # --- 2. 埋め込み情報の取得 ---
        index_key_raw = self.index_key_combo.get()
        index_key_to_embed = ""
        text_color = None  # fitzデフォルト (黒)

        if index_key_raw != "（未選択）":
            index_key_to_embed = index_key_raw
            key_colors_dict = self.parent_app.config_data.get("key_colors", {})
            hex_color = key_colors_dict.get(index_key_raw.lower())
            if hex_color:
                text_color = hex_to_rgb_tuple(hex_color)

        comment_to_embed = self.comment_textbox.get("1.0", "end-1c").strip()

        # 引用Keyリストを取得 ([[Key: Title]]形式に対応)
        cited_keys_str = self.cited_keys_entry.get("1.0", "end-1c").strip()
        cited_keys_list = []
        if cited_keys_str:
            # 1. Key (14桁以上の数字) を抽出するための正規表現
            key_regex = re.compile(r"(\d{14,})")
            # 2. カンマで分割 (複数のKey/リンクが入力された場合に対応)
            parts = re.split(r"[,\n]+", cited_keys_str)
            for part in parts:
                # 3. 各部分から Key (14桁以上の数字) を検索
                match = key_regex.search(part.strip())
                # 4. 見つかったKeyのみをリストに追加
                if match and match.group(1) not in cited_keys_list:
                    cited_keys_list.append(match.group(1))

        # SIST 02 書誌情報の構築
        # 1. UIから生のテキストを取得
        raw_sist_author = self.sist_author_entry.get().strip()
        raw_sist_title = self.sist_title_entry.get().strip()
        raw_sist_site = self.sist_site_entry.get().strip()
        raw_sist_date = self.sist_date_entry.get().strip()

        sist_view_date = datetime.datetime.now().strftime("%Y-%m-%d")
        sist_string_formal = None
        sist_string_readable = None

        # 2. 何か1つでも入力がある場合のみ、書誌情報文字列を構築する
        if raw_sist_author or raw_sist_title or raw_sist_site or raw_sist_date:

            # 3. 入力がある場合のみ、デフォルト値を適用
            sist_author = raw_sist_author or "（著者不明）"
            sist_title = raw_sist_title or "（タイトル不明）"
            sist_site = raw_sist_site or "（サイト名不明）"
            sist_date = raw_sist_date or "（更新日不明）"

            sist_string_formal = (
                f"{sist_author} . “{sist_title}” . "
                f"{sist_site} . {sist_date} . {url} , "
                f"(参照 {sist_view_date})"
            )
            sist_string_readable = (
                f"著者:{sist_author}\n\nページ名:\n“{sist_title}”\n\n"
                f"サイト名:\n{sist_site}\n\n"
                f"入手先:\n{url}\n\n更新日:{sist_date} (参照:{sist_view_date})"
            )

        # --- 3. 出力先フォルダを選択 ---
        dest_folder = filedialog.askdirectory(title="出力先フォルダを選択", parent=self)
        if not dest_folder:
            return
        dest_path = Path(dest_folder)

        # --- 4. 一時フォルダの準備 ---
        if self.temp_dir and self.temp_dir.exists():
            shutil.rmtree(self.temp_dir)
        self.temp_dir = Path(tempfile.mkdtemp(prefix="synapsen_clip_"))

        # 一時ファイルの「ベース」パスのみ定義
        temp_raw_item_path = self.temp_dir / f"{base_name}_raw_download"

        self.status_label.configure(text="PDF変換中...", text_color="gray")
        self.run_button.configure(state="disabled")
        self.fetch_button.configure(state="disabled")
        self.update_idletasks()

        conversion_success = False
        try:
            if (
                self.fetched_content_type
                and "text/html" not in self.fetched_content_type
            ):
                # ダウンロード処理
                parsed_url = urlparse(url)
                fname = unquote(Path(parsed_url.path).name)
                if not Path(fname).suffix:
                    if "pdf" in self.fetched_content_type:
                        fname += ".pdf"
                    elif "image" in self.fetched_content_type:
                        fname += ".jpg"
                    else:
                        fname += ".dat"
                temp_raw_item_path = temp_raw_item_path.with_name(fname)

                headers = {"User-Agent": "Mozilla/5.0"}
                req = urllib.request.Request(url, headers=headers)
                with urllib.request.urlopen(req, timeout=60) as response:
                    with open(temp_raw_item_path, "wb") as out_file:
                        shutil.copyfileobj(response, out_file)
                conversion_success = True
            else:
                # Playwright PDF化
                # HTMLなのにページがロードされていない場合 (エラー後など)
                if self.page is None:
                    raise Exception("ブラウザセッションがありません。")

                temp_raw_item_path = self.temp_dir / f"{base_name}_raw.pdf"
                paper_size_format = self.parent_app.config_data.get("paper_size", "A4")
                self.page.pdf(
                    path=str(temp_raw_item_path),
                    format=paper_size_format,
                    print_background=True,
                    margin={"top": "0", "bottom": "0", "left": "0", "right": "0"},
                )
                conversion_success = True

        except (
            PlaywrightError,
            PlaywrightTimeoutError,
            urllib.error.URLError,
            Exception,
        ) as e:
            # --- フォールバック: 簡易PDF生成 ---
            logger.error(f"Webクリップ変換失敗 (フォールバック実行): {e}")
            self.status_label.configure(
                text="取得失敗。書誌情報のみのPDFを生成します...", text_color="orange"
            )
            self.update_idletasks()

            try:
                temp_raw_item_path = self.temp_dir / f"{base_name}_fallback.pdf"

                doc = fitz.open()
                pw = self.parent_app.paper_width
                ph = self.parent_app.paper_height
                page = doc.new_page(width=pw, height=ph)

                font_name = "embed_font"
                if font_path and Path(font_path).is_file():
                    try:
                        page.insert_font(fontname=font_name, fontfile=font_path)
                    except Exception:
                        font_name = "helv"
                else:
                    font_name = "helv"

                fallback_text = (
                    f"【Webクリップ取得失敗】\n\n"
                    f"以下のURLからの取得に失敗しました (タイムアウトまたはエラー)。\n"
                    f"URL: {url}\n\n"
                    f"エラー内容: {str(e)}\n\n"
                    f"--------------------------------------------------\n"
                    f"■ 書誌情報 (手動入力/自動取得分)\n"
                    f"{sist_string_readable if sist_string_readable else '（情報なし）'}\n\n"
                    f"■ コメント\n"
                    f"{comment_to_embed}\n"
                )

                rect = fitz.Rect(50, 50, pw - 50, ph - 50)
                page.insert_textbox(
                    rect, fallback_text, fontsize=10, fontname=font_name
                )

                doc.save(str(temp_raw_item_path))
                doc.close()
                conversion_success = True

            except Exception as e_fallback:
                logger.error(f"フォールバックPDF生成も失敗: {e_fallback}")
                messagebox.showerror(
                    "エラー",
                    f"PDF作成に完全に失敗しました:\n{e}\n\n(簡易PDF生成エラー: {e_fallback})",
                    parent=self,
                )
                self.run_button.configure(state="normal")
                self.fetch_button.configure(state="normal")
                return

        if not conversion_success:
            return

        self.status_label.configure(text="正規化処理中...")
        self.update_idletasks()

        try:
            items_to_process = [(temp_raw_item_path, base_name)]
            self.parent_app.execute_normalization_process(items_to_process, dest_path)

            final_output_pdf = dest_path / f"{base_name}.pdf"
            config_data = self.parent_app.config_data
            key_rect_tuple = config_data.get("key_rect", (0, 0, 0, 0))
            refs_qr_size = config_data.get("refs_qr_size", 75)

            add_metadata_to_clip(
                pdf_path_str=str(final_output_pdf),
                font_path=font_path,
                paper_width=self.parent_app.paper_width,
                paper_height=self.parent_app.paper_height,
                key_rect_tuple=key_rect_tuple,
                index_key_to_embed=index_key_to_embed,
                text_color=text_color,
                comment_to_embed=comment_to_embed,
                sist_string_formal=sist_string_formal,
                sist_string_readable=sist_string_readable,
                base_name=base_name,
                cited_keys_list=cited_keys_list,
                refs_qr_size_pt=refs_qr_size,
                extra_keywords=["Synapsen:WebClip"],
            )
            self.on_close()

        except Exception as e:
            messagebox.showerror(
                "エラー", f"処理中にエラーが発生しました:\n{e}", parent=self
            )
            self.run_button.configure(state="normal")
            self.fetch_button.configure(state="normal")

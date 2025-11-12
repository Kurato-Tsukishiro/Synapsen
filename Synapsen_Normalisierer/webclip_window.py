import customtkinter as ctk
from tkinter import filedialog, messagebox
from pathlib import Path
import datetime
import tempfile
import shutil
import fitz  # PyMuPDF (情報埋め込み用)
from urllib.parse import urlparse  # サイト名取得用

from pdf_utils import add_metadata_to_web_clip, hex_to_rgb_tuple

# --- Playwright インポート ---
sync_playwright = None
PlaywrightError = Exception
PlaywrightTimeoutError = Exception

try:
    from playwright.sync_api import sync_playwright, Error, TimeoutError
    PlaywrightError = Error
    PlaywrightTimeoutError = TimeoutError
except ImportError:
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

        # --- Playwright (実行コンテキスト) ---
        self.playwright_context = None
        self.browser = None
        self.page = None

        # --- Playwrightがインストールされているかチェック ---
        if sync_playwright is None:
            messagebox.showerror(
                "ライブラリ不足エラー",
                "Webクリップ機能に必要な 'playwright' ライブラリが見つかりません。\n\n" +
                "コマンドプロンプトで以下のコマンドを実行してください:\n" +
                "1. pip install playwright\n" +
                "2. playwright install",
                parent=parent_app
            )
            # ウィンドウの表示シーケンスから抜け、即座に閉じる
            self.after(100, self.destroy)
            return

        self.title("Webクリップで正規化")
        self.geometry("450x670")

        if self.parent_app.icon_path:
            try:
                self.iconbitmap(default=str(self.parent_app.icon_path))
            except Exception as e:
                print(f"Icon set error (WebClip Window): {e}")

        self.protocol("WM_DELETE_WINDOW", self.on_close)
        self.transient(parent_app)
        self.grab_set()

        # --- [UI定義] ---

        # --- 1. URLとファイル名 ---
        input_frame = ctk.CTkFrame(self, fg_color="gray25")
        input_frame.pack(pady=10, padx=10, fill="x")

        ctk.CTkLabel(
            input_frame, text="URL:", width=80
            ).grid(row=0, column=0, padx=5, pady=5, sticky="w")
        self.url_entry = ctk.CTkEntry(
            input_frame, placeholder_text="https://...")
        self.url_entry.grid(
            row=0, column=1, columnspan=2, padx=5, pady=5, sticky="ew")

        ctk.CTkLabel(
            input_frame, text="ファイル名:", width=80
            ).grid(row=1, column=0, padx=5, pady=5, sticky="w")
        now = datetime.datetime.now()
        default_base_name = f"{now.strftime('%Y%m%d_%H%M%S')}_WebClip"
        self.filename_var = ctk.StringVar(value=default_base_name)
        self.filename_entry = ctk.CTkEntry(
            input_frame, textvariable=self.filename_var)
        self.filename_entry.grid(
            row=1, column=1, padx=5, pady=5, sticky="ew")
        ctk.CTkLabel(
            input_frame, text=".pdf"
            ).grid(row=1, column=2, padx=5, pady=5, sticky="w")

        input_frame.grid_columnconfigure(1, weight=1)

        # --- 2. ページ情報取得ボタン ---
        self.fetch_button = ctk.CTkButton(
            self,
            text="1. ページ情報取得 (書誌情報)",
            command=self.fetch_page_info
        )
        self.fetch_button.pack(pady=5, padx=10, fill="x")

        # --- 3. 書誌情報 (SIST 02) 編集フレーム ---
        ctk.CTkLabel(
            self, text="書誌情報 (SIST 02準拠)", anchor="w"
            ).pack(pady=(10, 0), padx=10, fill="x")
        sist_frame = ctk.CTkFrame(self)
        sist_frame.pack(pady=(0, 10), padx=10, fill="x")
        sist_frame.grid_columnconfigure(1, weight=1)

        ctk.CTkLabel(
            sist_frame, text="著者名:"
            ).grid(row=0, column=0, padx=5, pady=5, sticky="w")
        self.sist_author_entry = ctk.CTkEntry(
            sist_frame, placeholder_text="（自動取得試行）")
        self.sist_author_entry.grid(
            row=0, column=1, padx=5, pady=5, sticky="ew")

        ctk.CTkLabel(
            sist_frame, text="ページ名:"
            ).grid(row=1, column=0, padx=5, pady=5, sticky="w")
        self.sist_title_entry = ctk.CTkEntry(
            sist_frame, placeholder_text="（自動取得試行）")
        self.sist_title_entry.grid(
            row=1, column=1, padx=5, pady=5, sticky="ew")

        ctk.CTkLabel(
            sist_frame, text="サイト名:"
            ).grid(row=2, column=0, padx=5, pady=5, sticky="w")
        self.sist_site_entry = ctk.CTkEntry(
            sist_frame, placeholder_text="（自動取得試行）")
        self.sist_site_entry.grid(
            row=2, column=1, padx=5, pady=5, sticky="ew")

        ctk.CTkLabel(
            sist_frame, text="更新日:"
            ).grid(row=3, column=0, padx=5, pady=5, sticky="w")
        self.sist_date_entry = ctk.CTkEntry(
            sist_frame, placeholder_text="（自動取得試行, YYYY-MM-DD）")
        self.sist_date_entry.grid(
            row=3, column=1, padx=5, pady=5, sticky="ew")

        # --- 4. IndexKey 選択 ---
        ctk.CTkLabel(
            self, text="IndexKey (PDF 1ページ目に埋込)", anchor="w"
            ).pack(pady=(10, 0), padx=10, fill="x")
        key_frame = ctk.CTkFrame(self)
        key_frame.pack(pady=(0, 10), padx=10, fill="x")

        key_options = self.parent_app.config_data.get(
                'commonplace_keys_options', [])

        self.index_key_combo = ctk.CTkComboBox(
            key_frame,
            values=["（未選択）"] + key_options
        )
        self.index_key_combo.set("（未選択）")
        self.index_key_combo.pack(fill="x", padx=5, pady=5)

        # --- 5. コメント入力 ---
        ctk.CTkLabel(
            self, text="コメント (PDF 1ページ目に埋込)", anchor="w"
            ).pack(pady=(10, 0), padx=10, fill="x")
        comment_frame = ctk.CTkFrame(self)
        comment_frame.pack(pady=(0, 10), padx=10, fill="both", expand=True)
        self.comment_textbox = ctk.CTkTextbox(comment_frame, height=80)
        self.comment_textbox.pack(fill="both", expand=True, padx=5, pady=5)

        # --- 6. 実行ボタン ---
        self.run_button = ctk.CTkButton(
            self,
            text="2. 出力先を選んでクリップ実行",
            command=self.run_webclip_process,
            state="disabled"  # 初期状態は無効
        )
        self.run_button.pack(pady=10, padx=10, fill="x", ipady=10)

        # --- 7. ステータスラベル ---
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
                print(f"Playwright page close error: {e}")
        if self.browser:
            try:
                self.browser.close()
            except Exception as e:
                print(f"Playwright browser close error: {e}")
        if self.playwright_context:
            try:
                self.playwright_context.stop()
            except Exception as e:
                print(f"Playwright context stop error: {e}")

        # 一時フォルダのクリーンアップ
        if self.temp_dir and self.temp_dir.exists():
            try:
                shutil.rmtree(self.temp_dir)
            except Exception as e:
                print(f"WebClip一時フォルダの削除に失敗: {e}")

        self.grab_release()
        self.destroy()

    def fetch_page_info(self) -> None:
        """
        「1. ページ情報取得」ボタンの処理。
        Playwrightを使用してURLにアクセスし、書誌情報（タイトル、著者、サイト名、
        更新日）を抽出しようと試みます。
        タイムアウトやエラーが発生した場合は、限定的な情報を取得します。
        """
        url = self.url_entry.get().strip()
        if not url.startswith("http://") and not url.startswith("https://"):
            messagebox.showerror(
                "入力エラー", "有効なURL (https://...) を入力してください。", parent=self)
            return

        self.status_label.configure(
            text="ページ情報を取得中 (最大1分)...", text_color="gray")
        self.fetch_button.configure(state="disabled")
        self.run_button.configure(state="disabled")
        self.update_idletasks()

        try:
            # 1. Playwrightセッションがなければ開始する
            if self.playwright_context is None:
                self.status_label.configure(text="Playwrightを起動中...")
                self.update_idletasks()
                self.playwright_context = sync_playwright().start()
                self.browser = self.playwright_context.chromium.launch()
                self.page = self.browser.new_page()

            # 2. ページに移動 (タイムアウト1分)
            self.status_label.configure(text=f"ページに移動中:\n{url[:60]}...")
            self.update_idletasks()
            self.page.goto(
                url, wait_until='domcontentloaded', timeout=60000)

            # --- 3. 成功時のロジック (詳細取得) ---
            self.status_label.configure(
                text="ページ情報(詳細)を取得中...", text_color="gray")
            self.update_idletasks()

            self.page_title_cache = self.page.title()

            # メタデータ抽出 (複数のセレクタ候補を試行)
            author = self.page.locator(
                'meta[name="author"], '
                'meta[property="og:author"], '
                'meta[property="article:author"]'
            ).first.get_attribute("content")

            site_name = self.page.locator(
                'meta[property="og:site_name"]'
            ).first.get_attribute("content")

            if not site_name:
                self.site_name_cache = urlparse(url).netloc  # フォールバック
            else:
                self.site_name_cache = site_name

            date_str = self.page.locator(
                'meta[property="article:published_time"], '
                'meta[property="og:updated_time"], '
                'time[datetime]'
            ).first.get_attribute("datetime")

            iso_date = ""
            if date_str:
                try:
                    iso_date = date_str.split('T')[0]  # 'YYYY-MM-DD' の部分のみ取得
                except Exception:
                    iso_date = ""

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

            self.status_label.configure(text="ページ情報を取得しました。内容を確認・編集してください。")

        except (PlaywrightTimeoutError, PlaywrightError) as e:
            # 5. タイムアウト時のロジック (page.title と urlparse のみ取得)
            print(f"ページ読み込みがタイムアウトしました: {e}")
            self.status_label.configure(
                text="タイムアウト。簡易情報(タイトル/ドメイン)を取得します...",
                text_color="orange"
            )
            self.update_idletasks()

            try:
                self.page_title_cache = self.page.title()
                parsed_url = urlparse(url)
                self.site_name_cache = parsed_url.netloc

                self.sist_title_entry.delete(0, "end")
                self.sist_title_entry.insert(0, self.page_title_cache)
                self.sist_site_entry.delete(0, "end")
                self.sist_site_entry.insert(0, self.site_name_cache)
                self.sist_author_entry.delete(0, "end")
                self.sist_date_entry.delete(0, "end")

                messagebox.showwarning(
                    "情報取得タイムアウト",
                    f"ページの読み込みがタイムアウトしました (1分)。\n{e}\n\n"
                    "タイトルとサイト名のみ取得を試みました。\n"
                    "著者名や更新日は手動で入力してください。",
                    parent=self
                )
            except Exception as simple_e:
                messagebox.showerror(
                    "情報取得エラー",
                    "タイムアウト後の簡易情報取得にも失敗しました:\n"
                    f"{simple_e}",
                    parent=self
                )
                self.status_label.configure(
                    text="情報取得失敗（タイムアウト）。手動で入力してください。",
                    text_color="orange"
                )
                self.page_title_cache = ""
                self.site_name_cache = ""

        except Exception as e:
            messagebox.showerror(
                "情報取得エラー", f"予期せぬエラーが発生しました:\n{e}", parent=self)
            self.status_label.configure(
                text="情報取得エラー。", text_color="orange")

        finally:
            self.fetch_button.configure(state="normal")
            self.run_button.configure(state="normal")  # 実行ボタンを有効化
            self.bell()

    def run_webclip_process(self) -> None:
        """
        「2. クリップ実行」ボタンの処理。

        PlaywrightでPDF化し、PyMuPDFで情報（IndexKey, 書誌情報, コメント）を
        1ページ目と最終ページに埋め込み、
        親アプリの `execute_normalization_process` を呼び出して
        最終的な正規化を行います。
        """
        url = self.url_entry.get().strip()
        base_name = self.filename_var.get().strip()

        # --- 1. 入力バリデーション ---
        if not url.startswith("http://") and not url.startswith("https://"):
            messagebox.showerror(
                "入力エラー", "有効なURL (https://...) を入力してください。", parent=self)
            return
        if not base_name:
            messagebox.showerror(
                "入力エラー", "ファイル名を入力してください。", parent=self)
            return
        font_path = self.parent_app.font_path
        if (not font_path or not Path(font_path).is_file()):
            self.parent_app.status_label.configure(
                text="エラー: config.iniで有効なフォントパスが指定されていません。",
                text_color="orange"
            )
            return
        if self.page is None:
            messagebox.showerror(
                "エラー", "先に「1. ページ情報取得」ボタンを押して、ページを読み込んでください。", parent=self)
            return

        # --- 2. 埋め込み情報の取得 ---
        index_key_raw = self.index_key_combo.get()
        index_key_to_embed = ""
        text_color = None  # fitzデフォルト (黒)

        if index_key_raw != "（未選択）":
            index_key_to_embed = index_key_raw
            key_colors_dict = self.parent_app.config_data.get('key_colors', {})
            hex_color = key_colors_dict.get(index_key_raw.lower())
            if hex_color:
                text_color = hex_to_rgb_tuple(hex_color)

        comment_to_embed = self.comment_textbox.get("1.0", "end-1c").strip()

        # SIST 02 書誌情報の構築
        # 1. UIから生のテキストを取得
        raw_sist_author = self.sist_author_entry.get().strip()
        raw_sist_title = self.sist_title_entry.get().strip()
        raw_sist_site = self.sist_site_entry.get().strip()
        raw_sist_date = self.sist_date_entry.get().strip()

        sist_view_date = datetime.datetime.now().strftime('%Y-%m-%d')
        sist_string_formal = None
        sist_string_readable = None

        # 2. 何か1つでも入力がある場合のみ、書誌情報文字列を構築する
        if (raw_sist_author or raw_sist_title or
                raw_sist_site or raw_sist_date):

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
        dest_folder = filedialog.askdirectory(
                comment_to_embed=comment_to_embed,
                sist_string_formal=sist_string_formal,
                sist_string_readable=sist_string_readable
            )
        if not dest_folder:
            return
        dest_path = Path(dest_folder)

        # --- 4. 一時フォルダの準備 ---
        if self.temp_dir and self.temp_dir.exists():
            shutil.rmtree(self.temp_dir)
        self.temp_dir = Path(tempfile.mkdtemp(prefix="synapsen_clip_"))
        temp_pdf_path = self.temp_dir / f"{base_name}_raw.pdf"

        self.status_label.configure(
            text="WebページをPDFに変換中...", text_color="gray")
        self.run_button.configure(state="disabled")
        self.fetch_button.configure(state="disabled")
        self.update_idletasks()

        # --- 5. PlaywrightによるPDF化 ---
        try:
            # Playwrightの用紙サイズは親アプリ(A4/A5)に合わせる
            paper_size_format = self.parent_app.config_data.get(
                'paper_size', 'A4')

            self.page.pdf(
                path=str(temp_pdf_path), format=paper_size_format,  # [変更]
                print_background=True,
                margin={
                    'top': '1cm', 'bottom': '1cm',
                    'left': '1cm', 'right': '1cm'
                }
            )
        except Exception as pdf_e:
            # 印刷失敗時のフォールバック (簡易PDF生成)
            print(f"PDFの印刷に失敗: {pdf_e}")
            self.status_label.configure(text="印刷失敗。最小限の簡易PDFを生成します...")
            self.update_idletasks()
            try:
                doc = fitz.open()
                # 親アプリの用紙サイズを使用
                paper_width = self.parent_app.paper_width
                paper_height = self.parent_app.paper_height
                pdf_page = doc.new_page(width=paper_width, height=paper_height)

                text_to_insert = (
                    f"簡易Webクリップ (タイムアウト・印刷失敗)\n\n"
                    f"書誌情報(SIST 02):\n{sist_string_formal}\n\n"
                    f"コメント:\n{comment_to_embed}\n\n"
                )
                rect = fitz.Rect(
                    50, 100,
                    pdf_page.rect.width - 50, pdf_page.rect.height - 50
                )
                pdf_page.insert_textbox(rect, text_to_insert,
                                        fontsize=10, fontname="helv", align=0)
                doc.save(str(temp_pdf_path))
                doc.close()
            except Exception as e:
                messagebox.showerror(
                    "Webクリップエラー", f"簡易PDFの生成にも失敗しました:\n{e}", parent=self)
                self.status_label.configure(
                    text="エラーが発生しました。", text_color="orange")
                self.run_button.configure(state="normal")
                self.fetch_button.configure(state="normal")
                return

        # --- 6. PDFへの情報埋め込み (pdf_utils 経由 PyMuPDF) ---
        self.status_label.configure(text="PDFに情報を埋め込み中...")
        self.update_idletasks()
        try:
            # 親アプリから設定を取得
            config_data = self.parent_app.config_data
            key_rect_tuple = config_data.get('key_rect', (0, 0, 0, 0))
            paper_width = self.parent_app.paper_width
            paper_height = self.parent_app.paper_height

            # ヘルパー関数 (add_metadata_to_web_clip) を呼び出し
            add_metadata_to_web_clip(
                pdf_path_str=str(temp_pdf_path),
                font_path=font_path,
                paper_width=paper_width,
                paper_height=paper_height,
                key_rect_tuple=key_rect_tuple,
                index_key_to_embed=index_key_to_embed,
                text_color=text_color,
                comment_to_embed=comment_to_embed,
                sist_string_formal=sist_string_formal,
                sist_string_readable=sist_string_readable
            )

        except Exception as e:
            messagebox.showerror(
                "情報埋め込みエラー", f"PDFへの情報埋め込みに失敗しました:\n{e}", parent=self)
            # エラーが起きても、PDF化自体は成功しているので処理は続行

        # --- 7. 親アプリの正規化処理を呼び出す ---
        self.status_label.configure(text="PDF正規化処理を実行中...")
        self.update_idletasks()

        # execute_normalization_process は (Path, base_name) のタプルリストを期待する
        items_to_process = [(temp_pdf_path, base_name)]

        try:
            self.parent_app.execute_normalization_process(
                items_to_process, dest_path
            )
            self.on_close()  # 成功したらウィンドウを閉じる

        except Exception as e:
            messagebox.showerror(
                "正規化処理エラー", f"処理中にエラーが発生しました:\n{e}", parent=self)
            self.status_label.configure(
                text="エラーが発生しました。", text_color="orange")
            self.run_button.configure(state="normal")
            self.fetch_button.configure(state="normal")

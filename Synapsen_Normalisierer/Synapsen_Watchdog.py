import sys
import time
import logging
import configparser
import threading
import customtkinter as ctk
from pathlib import Path
from watchdog.observers import Observer
from watchdog.events import FileSystemEventHandler

# --- パス解決ロジック ---
if getattr(sys, "frozen", False):
    ROOT_DIR = Path(sys.executable).parent
    CURRENT_DIR = Path(__file__).parent
else:
    CURRENT_DIR = Path(__file__).parent
    ROOT_DIR = CURRENT_DIR.parent
sys.path.append(str(CURRENT_DIR))

current_dir = Path(__file__).parent
root_dir = current_dir.parent
if str(root_dir) not in sys.path:
    sys.path.append(str(root_dir))

try:
    from pdf_utils import normalize_pdf_to_papersize
except ImportError:
    pass  # GUI内でエラー表示するためここではパス

from theme import SemanticColors as Colors  # noqa: E402


# --- ロギング設定 (GUI出力用) ---
class TextHandler(logging.Handler):
    """ログをGUIのテキストボックスに出力するためのハンドラ"""

    def __init__(self, text_widget):
        super().__init__()
        self.text_widget = text_widget

    def emit(self, record):
        msg = self.format(record)

        def append():
            self.text_widget.configure(state="normal")
            self.text_widget.insert("end", msg + "\n")
            self.text_widget.see("end")
            self.text_widget.configure(state="disabled")

        # GUIスレッドで実行
        self.text_widget.after(0, append)


# --- 設定読み込み ---
def load_watch_config():
    config_path = ROOT_DIR / "config.ini"
    config = configparser.ConfigParser()

    if not config_path.exists():
        return None, None, "A4", f"Config file not found: {config_path}"

    try:
        config.read(config_path, encoding="utf-8")
        if "Watchdog" not in config:
            return None, None, "A4", "[Watchdog] section not found in config.ini"

        watch_dir = config["Watchdog"].get("watch_dir", "")
        output_dir = config["Watchdog"].get("output_dir", "")
        target_size = config["Watchdog"].get("target_size", "A4")

        return watch_dir, output_dir, target_size, None

    except Exception as e:
        return None, None, "A4", f"Config load error: {e}"


PAPER_SIZES = {"A4": (595.276, 841.89), "A5": (419.528, 595.276)}


# --- イベントハンドラ ---
class PDFHandler(FileSystemEventHandler):
    def __init__(self, output_dir, target_size, logger):
        self.output_dir = Path(output_dir)
        self.target_size = target_size
        self.logger = logger
        self.processing_files = set()

    def on_created(self, event):
        if not event.is_directory:
            self.process_file(Path(event.src_path))

    def on_moved(self, event):
        if not event.is_directory:
            self.process_file(Path(event.dest_path))

    def process_file(self, filepath: Path):
        """
        単一ファイルの処理ロジック。
        イベントハンドラからも、初期スキャンからも呼ばれる。
        """
        # PDF以外、一時ファイル、隠しファイルは無視
        if filepath.suffix.lower() != ".pdf":
            return
        if filepath.name.startswith("~") or filepath.name.startswith("."):
            return

        # 既に処理中ならスキップ
        if filepath in self.processing_files:
            return

        # 出力先に同名ファイルが存在するかチェック
        output_path = self.output_dir / filepath.name
        if output_path.exists():
            # 既に存在する場合、何もしない（ログも出さない方が静かで良い）
            # ただし、入力ファイルの更新日時が出力より新しい場合は再処理するロジックも考えられますが、
            # シンプルな運用なら「存在すればスキップ」で十分です。
            return

        self.processing_files.add(filepath)
        threading.Thread(
            target=self._run_conversion, args=(filepath, output_path), daemon=True
        ).start()

    def _run_conversion(self, filepath, output_path):
        try:
            self.logger.info(f"検知: {filepath.name}")

            # ファイル書き込み完了待ち
            if not self._wait_for_file_ready(filepath):
                self.logger.warning(f"タイムアウト (アクセス不可): {filepath.name}")
                return

            self.logger.info(f"変換開始 ({self.target_size}): {filepath.name}")
            paper_width, paper_height = PAPER_SIZES.get(
                self.target_size, PAPER_SIZES["A4"]
            )

            try:
                normalize_pdf_to_papersize(
                    str(filepath),
                    str(output_path),
                    paper_width,
                    paper_height,
                    target_format=self.target_size,
                )
                self.logger.info(f"完了 -> {output_path.name}")

                try:
                    filepath.unlink()
                    self.logger.info("元ファイルを削除しました。")
                except FileNotFoundError:
                    # ファイルが既にない場合は「削除済み」とみなしてエラーにしない
                    self.logger.info("元ファイルは既に削除されています。")
                except PermissionError:
                    # 使用中で消せない場合はログだけ出してスルー（次回の手動削除に委ねる）
                    self.logger.warning("削除失敗: ファイルが使用中です。")
                except Exception as e:
                    self.logger.error(f"削除失敗: {e}")

            except Exception as e:
                self.logger.error(f"変換失敗: {e}")

        except Exception as e:
            self.logger.error(f"予期せぬエラー: {e}")
        finally:
            self.processing_files.discard(filepath)

    def _wait_for_file_ready(self, filepath, timeout=30, stable_duration=2):
        start_time = time.time()
        last_size = -1
        stable_start = None

        while True:
            if time.time() - start_time > timeout:
                return False

            try:
                if not filepath.exists():
                    return False

                size = filepath.stat().st_size

                if size == last_size and size > 0:
                    if stable_start is None:
                        stable_start = time.time()
                    elif time.time() - stable_start >= stable_duration:
                        return True
                else:
                    last_size = size
                    stable_start = None
            except Exception:
                pass
            time.sleep(1)


# --- GUI アプリケーション ---
class WatchdogApp(ctk.CTk):
    def __init__(self):
        super().__init__()
        self.title("Synapsen Watchdog")
        self.geometry("500x400")
        self.configure(
            fg_color=(Colors.BACKGROUND_HOLLOW, Colors.BACKGROUND_DARK_HOLLOW)
        )

        # アイコン設定
        icon_path = ROOT_DIR / "assets" / "synapsen.ico"
        if icon_path.exists():
            try:
                self.iconbitmap(default=str(icon_path))
            except Exception:
                pass

        # テーマカラー (蘇芳色)
        self.theme_color = Colors.WATCHDOG

        self.observer = None
        self._setup_ui()
        self._start_watchdog()

    def _setup_ui(self):
        self.grid_columnconfigure(0, weight=1)
        self.grid_rowconfigure(1, weight=1)

        # ヘッダー
        header_frame = ctk.CTkFrame(self, fg_color=self.theme_color, corner_radius=0)
        header_frame.grid(row=0, column=0, sticky="ew")

        title_label = ctk.CTkLabel(
            header_frame,
            text="自動正規化 監視中",
            font=("Arial", 18, "bold"),
            text_color="white",
        )
        title_label.pack(pady=10)

        # ログ表示エリア
        self.log_textbox = ctk.CTkTextbox(
            self,
            font=("Consolas", 12),
            fg_color=(Colors.BACKGROUND_HOLLOW, Colors.BACKGROUND_DARK_HOLLOW),
        )
        self.log_textbox.grid(row=1, column=0, padx=10, pady=10, sticky="nsew")
        self.log_textbox.configure(state="disabled")

        # ログハンドラの設定
        self.logger = logging.getLogger("WatchdogGUI")
        self.logger.setLevel(logging.INFO)
        # 既存ハンドラ削除
        if self.logger.handlers:
            self.logger.handlers = []

        handler = TextHandler(self.log_textbox)
        formatter = logging.Formatter("%(asctime)s - %(message)s", datefmt="%H:%M:%S")
        handler.setFormatter(formatter)
        self.logger.addHandler(handler)

        # フッター
        footer_frame = ctk.CTkFrame(self, fg_color="transparent")
        footer_frame.grid(row=2, column=0, padx=10, pady=10, sticky="ew")

        self.info_label = ctk.CTkLabel(
            footer_frame, text="初期化中...", text_color="gray"
        )
        self.info_label.pack(side="left")

        stop_btn = ctk.CTkButton(
            footer_frame,
            text="停止して終了",
            fg_color=Colors.adjust_brightness(Colors.BACKGROUND_PANEL, 0.6),
            hover_color=Colors.adjust_brightness(Colors.BACKGROUND_PANEL, 0.4),
            width=100,
            command=self.on_closing,
        )
        stop_btn.pack(side="right")

        self.protocol("WM_DELETE_WINDOW", self.on_closing)

    def _start_watchdog(self):
        watch_dir, output_dir, target_size, error_msg = load_watch_config()

        if error_msg:
            self.logger.error(error_msg)
            self.info_label.configure(text="設定エラー")
            return

        w_path = Path(watch_dir)
        o_path = Path(output_dir)

        if not w_path.exists():
            self.logger.error(f"監視フォルダが見つかりません: {w_path}")
            return

        if not o_path.exists():
            try:
                o_path.mkdir(parents=True, exist_ok=True)
            except Exception as e:
                self.logger.error(f"出力フォルダ作成エラー: {e}")
                return

        self.logger.info(f"監視開始: {w_path}")
        self.logger.info(f"出力先: {o_path} ({target_size})")
        self.info_label.configure(text=f"監視中: {w_path.name}")

        event_handler = PDFHandler(o_path, target_size, self.logger)
        self.observer = Observer()
        self.observer.schedule(event_handler, str(w_path), recursive=False)
        self.observer.start()

        # 初期スキャン (別スレッド)
        threading.Thread(
            target=self._initial_scan, args=(w_path, event_handler), daemon=True
        ).start()

    def _initial_scan(self, watch_path, handler):
        time.sleep(1)  # GUI表示待ち
        files = list(watch_path.glob("*.pdf"))
        if files:
            self.logger.info(f"既存ファイル {len(files)} 件を検出。処理を開始します...")
            for f in files:
                handler.process_file(f)

    def on_closing(self):
        if self.observer:
            self.logger.info("監視を停止しています...")
            self.observer.stop()
            # joinは時間がかかる場合があるため、UIを先に閉じる
        self.destroy()
        sys.exit(0)


def main():
    ctk.set_appearance_mode("System")
    app = WatchdogApp()
    app.mainloop()


if __name__ == "__main__":
    main()

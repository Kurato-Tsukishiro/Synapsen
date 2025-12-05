import sys
import time
import logging
import configparser
from pathlib import Path
from watchdog.observers import Observer
from watchdog.events import FileSystemEventHandler

# ルートディレクトリのパス解決
current_dir = Path(__file__).parent
root_dir = current_dir.parent
sys.path.append(str(current_dir))

try:
    from pdf_utils import normalize_pdf_to_papersize
except ImportError:
    print("Error: pdf_utils.py not found.")
    sys.exit(1)

# --------------------------------------------------------------------------
# ロギング設定
# --------------------------------------------------------------------------
logging.basicConfig(
    level=logging.INFO, format="%(asctime)s - %(message)s", datefmt="%Y-%m-%d %H:%M:%S"
)
logger = logging.getLogger("Watchdog")


# --------------------------------------------------------------------------
# 設定読み込み
# --------------------------------------------------------------------------
def load_watch_config():
    config_path = root_dir / "config.ini"
    config = configparser.ConfigParser()

    if not config_path.exists():
        logger.error(f"Config file not found: {config_path}")
        return None, None, "A4"

    try:
        config.read(config_path, encoding="utf-8")
        if "Watchdog" not in config:
            logger.warning(
                "[Watchdog] section not found in config.ini. Using defaults/empty."
            )
            return None, None, "A4"

        watch_dir = config["Watchdog"].get("watch_dir", "")
        output_dir = config["Watchdog"].get("output_dir", "")
        target_size = config["Watchdog"].get("target_size", "A4")

        return watch_dir, output_dir, target_size

    except Exception as e:
        logger.error(f"Config load error: {e}")
        return None, None, "A4"


# 用紙サイズの定義
PAPER_SIZES = {"A4": (595.276, 841.89), "A5": (419.528, 595.276)}


# --------------------------------------------------------------------------
# イベントハンドラ
# --------------------------------------------------------------------------
class PDFHandler(FileSystemEventHandler):
    def __init__(self, output_dir, target_size):
        self.output_dir = Path(output_dir)
        self.target_size = target_size
        self.processing_files = set()

    def on_created(self, event):
        self._process_event(event)

    def on_moved(self, event):
        class MockEvent:
            src_path = event.dest_path
            is_directory = event.is_directory

        self._process_event(MockEvent)

    def _process_event(self, event):
        """Watchdogイベントからのエントリポイント"""
        if event.is_directory:
            return

        filepath = Path(event.src_path)
        self.process_file(filepath)

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
        try:
            logger.info(f"検知: {filepath.name}")

            # 書き込み完了待機
            if self._wait_for_file_ready(filepath):
                logger.info(f"変換開始 ({self.target_size}): {filepath.name}")

                paper_width, paper_height = PAPER_SIZES.get(
                    self.target_size, PAPER_SIZES["A4"]
                )

                try:
                    # 正規化処理
                    normalize_pdf_to_papersize(
                        str(filepath),
                        str(output_path),
                        paper_width,
                        paper_height,
                        target_format=self.target_size,
                    )
                    logger.info(f"完了 -> {output_path}")

                    # 処理成功後に元ファイルを削除
                    try:
                        filepath.unlink()
                        logger.info(f"元ファイルを削除しました: {filepath.name}")
                    except Exception as e:
                        logger.error(f"元ファイルの削除に失敗: {e}")

                except Exception as e:
                    logger.error(f"変換失敗: {filepath.name}, Error: {e}")

            else:
                logger.warning(f"ファイルアクセス不可 (タイムアウト): {filepath.name}")

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

                current_size = filepath.stat().st_size

                if current_size == last_size and current_size > 0:
                    if stable_start is None:
                        stable_start = time.time()
                    elif time.time() - stable_start >= stable_duration:
                        return True
                else:
                    last_size = current_size
                    stable_start = None

            except PermissionError:
                stable_start = None
                pass
            except Exception as e:
                logger.debug(f"Wait error: {e}")

            time.sleep(1)


# --------------------------------------------------------------------------
# メイン処理
# --------------------------------------------------------------------------
def main():
    watch_dir, output_dir, target_size = load_watch_config()

    if not watch_dir or not output_dir:
        logger.error("config.ini 設定エラー")
        return

    watch_path = Path(watch_dir)
    output_path = Path(output_dir)

    if not watch_path.exists():
        logger.error(f"監視フォルダ不在: {watch_path}")
        return

    if not output_path.exists():
        try:
            output_path.mkdir(parents=True, exist_ok=True)
        except Exception as e:
            logger.error(f"出力フォルダ作成不可: {e}")
            return

    if watch_path.resolve() == output_path.resolve():
        logger.error("監視フォルダと出力フォルダを別にしてください。")
        return

    event_handler = PDFHandler(output_path, target_size)

    # --- 初期スキャン (起動時に既存ファイルを処理) ---
    logger.info("初期スキャンを実行中...")
    existing_files = list(watch_path.glob("*.pdf"))
    if existing_files:
        logger.info(f"{len(existing_files)} 件のファイルを検出しました。")
        for f in existing_files:
            event_handler.process_file(f)
    else:
        logger.info("処理対象ファイルはありません。")
    # ----------------------------------------------------

    observer = Observer()
    observer.schedule(event_handler, str(watch_path), recursive=False)

    observer.start()

    print("=" * 60)
    print(" Synapsen Auto-Normalizer Watchdog Running")
    print(f" 監視元: {watch_path}")
    print(f" 出力先: {output_path}")
    print(f" サイズ: {target_size} (縦)")
    print(" Ctrl+C で停止")
    print("=" * 60)

    try:
        while True:
            time.sleep(1)
    except KeyboardInterrupt:
        observer.stop()
        logger.info("監視を停止しました。")

    observer.join()


if __name__ == "__main__":
    main()

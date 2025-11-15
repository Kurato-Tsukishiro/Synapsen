import logging
import logging.handlers
import sys
from pathlib import Path

# logger.debug("..."):
#        開発中の詳細情報（変数の値など）。通常運用時は出力されません。
#        表示するには setLevel(logging.INFO) を logger.setLevel(logging.DEBUG) に変更します。
# logger.info("..."):
#        正常な動作の記録（「処理を開始しました」「完了しました」など）。
# logger.warning("..."):
#        処理は続行できるが、注意が必要なこと（「フォントが見つかりません、代替を使用します」など）。
# logger.error("..."):
#        処理が失敗したこと（例外発生など）。


class SensitiveDataFilter(logging.Filter):
    """
    機密情報フラグ(extra={'sensitive': True})を持つログをフィルタリングする。

    ロガーのレベルが INFO (またはそれ以上) に設定されている場合、
    機密ログの内容をマスクしたメッセージに置き換える。
    ロガーのレベルが DEBUG に設定されている場合のみ、元のメッセージを通過させる。
    """
    def filter(self, record):
        # 'sensitive' フラグが付いているログかチェック
        if getattr(record, 'sensitive', False):

            # ルートロガーの有効レベルを取得
            # (注: record.levelno ではなく、ロガー全体の *設定* を見る)
            effective_level = logging.getLogger().getEffectiveLevel()

            # もしロガー設定が DEBUG レベル「より上」(INFO, WARNING等) なら...
            if effective_level > logging.DEBUG:
                # ログメッセージをマスクした内容に上書きする
                record.msg = (
                    f"[SENSITIVE LOG: {record.levelname}] "
                    f"機密情報を含むログが {record.name}.py (L{record.lineno}) "
                    "で発生しました。"
                )
                # メッセージフォーマット(%sなど)を使っている場合に備え、引数を空にする
                record.args = ()

        # ログを通過させる
        return True


def setup_logging(app_name: str, log_folder_name: str = "logs"):
    """
    アプリケーションのロギング設定を初期化する。

    Args:
        app_name (str): ログファイル名に使用するアプリケーション名。
        log_folder_name (str): ログファイルを保存するフォルダ名。
    """
    # 1. ログ保存先パスの決定
    if getattr(sys, 'frozen', False):
        # .exeの場合: exeと同じ階層
        base_path = Path(sys.executable).parent
    else:
        # スクリプトの場合: このファイルの親ディレクトリ (プロジェクトルート)
        base_path = Path(__file__).parent

    log_dir = base_path / log_folder_name
    log_dir.mkdir(exist_ok=True)

    # ログファイルパス (例: logs/Synapsen_Nexus.log)
    # ★ ベースとなるファイル名を指定 (日付はハンドラが自動付与)
    log_file_path = log_dir / f"{app_name}.log"

    # 2. ルートロガーの設定
    logger = logging.getLogger()
    logger.setLevel(logging.INFO)

    # 既存のハンドラがあればクリア（リロード時の重複防止）
    if logger.handlers:
        logger.handlers.clear()

    # 3. フォーマッター作成
    # 出力例: 2023-10-27 10:00:00,123 - Synapsen_Nexus - INFO - 処理を開始します
    formatter = logging.Formatter(
        fmt='%(asctime)s - %(name)s - %(levelname)s - %(message)s',
        datefmt='%Y-%m-%d %H:%M:%S'
    )

    # フィルターのインスタンスを作成
    sensitive_filter = SensitiveDataFilter()

    # 4. ハンドラの設定

    # A. ファイルハンドラ
    # 毎日深夜0時にローテーションし、古いログを5世代までバックアップ
    file_handler = logging.handlers.TimedRotatingFileHandler(
        log_file_path,
        when='midnight',  # 毎日深夜 (日付が変わった瞬間)
        interval=1,       # 1日ごと
        backupCount=5,    # 5日分のバックアップを保持 (古いものは自動削除)
        encoding='utf-8'
    )
    # ローテーション後のファイル名 (例: Synapsen_Nexus.log.2025-11-16)
    file_handler.suffix = "%Y-%m-%d"

    file_handler.setFormatter(formatter)
    file_handler.setLevel(logging.INFO)
    file_handler.addFilter(sensitive_filter)  # フィルターをアタッチ
    logger.addHandler(file_handler)

    # B. コンソールハンドラ (標準出力)
    console_handler = logging.StreamHandler(sys.stdout)
    console_handler.setFormatter(formatter)
    console_handler.setLevel(logging.INFO)
    console_handler.addFilter(sensitive_filter)  # フィルターをアタッチ
    logger.addHandler(console_handler)

    # 5. 未処理例外のフック (重要)
    # GUIアプリで「何も言わずに落ちる」を防ぐため、エラーをログに記録する
    def handle_exception(exc_type, exc_value, exc_traceback):
        if issubclass(exc_type, KeyboardInterrupt):
            sys.__excepthook__(exc_type, exc_value, exc_traceback)
            return
        logger.error(
            "Uncaught exception",
            exc_info=(exc_type, exc_value, exc_traceback)
        )

    sys.excepthook = handle_exception

    logger.info(f"--- Logging Setup Complete: {app_name} ---")
    return logger

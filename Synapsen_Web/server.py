import sys
import os
import sqlite3
import logging
import signal  # 終了シグナル制御用
from functools import wraps
from pathlib import Path
from flask import (
    Flask,
    render_template,
    request,
    send_file,
    url_for,
    redirect,
    Response,
    send_from_directory,
)

# --- パスの設定 ---
# 親ディレクトリ(Synapsenルート)をパスに追加して Nexus のモジュールを読み込めるようにする
current_dir = Path(__file__).resolve().parent
root_dir = current_dir.parent
if str(root_dir) not in sys.path:
    sys.path.append(str(root_dir))

# Nexusの既存ロジックをインポート
from Synapsen_Nexus.utils import (  # noqa: E402
    load_app_config,
    fetch_notes_from_db,
    find_file_in_paths,
    count_notes_from_db,
    get_all_tags_with_count,
)
from Synapsen_Nexus.search_parser import parse_query_to_sql  # noqa: E402


def setup_console_logging():
    # 1. ルートロガー(全体)の設定
    root_logger = logging.getLogger()
    root_logger.setLevel(logging.INFO)

    # 既存のハンドラを全て削除（重複防止）
    if root_logger.handlers:
        root_logger.handlers.clear()

    # 標準出力(stdout)へのハンドラを作成
    # EXEのコンソール(CONOUT$)に確実に出力されるよう、sys.stdoutを明示的に指定
    handler = logging.StreamHandler(sys.stdout)
    handler.setLevel(logging.INFO)
    formatter = logging.Formatter(
        "%(asctime)s - %(name)s - %(levelname)s - %(message)s"
    )
    handler.setFormatter(formatter)

    # ハンドラはルートロガーに「1つだけ」追加する
    root_logger.addHandler(handler)

    # 2. Werkzeug (Flask通信ログ) の設定
    werkzeug_logger = logging.getLogger("werkzeug")
    werkzeug_logger.setLevel(logging.INFO)

    # Werkzeugが勝手に持っているハンドラがあれば削除する
    if werkzeug_logger.handlers:
        werkzeug_logger.handlers.clear()

    # 【重要】ハンドラは追加せず、親(ルートロガー)に任せる
    # propagate(伝播)をTrueにすることで、上のroot_logger.addHandler(handler)が処理してくれる
    werkzeug_logger.propagate = True

    return logging.getLogger("Synapsen_Web")


# ロガー初期化
logger = setup_console_logging()

# --- Flaskアプリ初期化 ---
if getattr(sys, "frozen", False):
    template_folder = Path(sys._MEIPASS) / "Synapsen_Web" / "templates"
    app = Flask(__name__, template_folder=str(template_folder))
else:
    app = Flask(__name__)

# --- 設定の読み込み ---
try:
    if getattr(sys, "frozen", False):
        base_path = Path(sys.executable).parent
    else:
        base_path = Path(__file__).parent

    config_data = load_app_config(base_path)
    DB_PATH = config_data["database_path"]
    PDF_ROOT = config_data["pdf_root_folder"]
    PDF_ARCHIVE = config_data["pdf_archive_folder"]

except Exception as e:
    logger.error(f"設定の読み込みに失敗しました: {e}")
    sys.exit(1)

# --- 認証設定 ---
USERNAME = "admin"
PASSWORD = "password"


def check_auth(username, password):
    """パスワード確認"""
    return username == USERNAME and password == PASSWORD


def authenticate():
    """認証エラー時のレスポンス"""
    return Response(
        "ログインが必要です。\nLogin Required",
        401,
        {"WWW-Authenticate": 'Basic realm="Login Required"'},
    )


def requires_auth(f):
    """認証デコレータ"""

    @wraps(f)
    def decorated(*args, **kwargs):
        auth = request.authorization
        if not auth or not check_auth(auth.username, auth.password):
            return authenticate()
        return f(*args, **kwargs)

    return decorated


def get_db_connection():
    conn = sqlite3.connect(f"file:{DB_PATH}?mode=ro", uri=True)
    conn.row_factory = sqlite3.Row
    return conn


# --- テンプレートフィルター ---
@app.template_filter("nl2br")
def nl2br_filter(s):
    """改行を<br>に変換"""
    return s.replace("\n", "<br>") if s else ""


@app.template_filter("linkify_key")
def linkify_key_filter(s):
    """[[Key]] 形式のリンクを <a href="..."> に変換"""
    if not s:
        return ""
    import re

    # [[Key]] または [[Key: Title]]
    def replacer(match):
        content = match.group(1)
        key = content.split(":")[0].strip()
        return f'<a href="{url_for("view_note", key=key)}">[[{content}]]</a>'

    return re.sub(r"\[\[(.*?)\]\]", replacer, s)


# --- ルーティング ---


@app.route("/favicon.ico")
def favicon():
    """ブラウザが自動要求するアイコンファイルを返す"""
    if getattr(sys, "frozen", False):
        # EXE化されている場合
        static_dir = Path(sys._MEIPASS) / "Synapsen_Web" / "static"
    else:
        # スクリプト実行の場合
        static_dir = Path(__file__).parent / "static"

    return send_from_directory(
        static_dir, "favicon.ico", mimetype="image/vnd.microsoft.icon"
    )


@app.route("/")
@requires_auth
def index():
    """ホーム画面（検索 + 新着リスト）"""
    query = request.args.get("q", "").strip()
    page = request.args.get("page", 0, type=int)
    limit = 20

    conn = get_db_connection()
    try:
        where_clause = ""
        params = []
        if query:
            where_clause, params = parse_query_to_sql(query, include_full_text=True)

        offset = page * limit
        notes = fetch_notes_from_db(
            conn, where_clause, params, limit=limit, offset=offset, sort_ascending=False
        )
        total_count = count_notes_from_db(conn, where_clause, params)
    finally:
        conn.close()

    max_page = max(0, (total_count - 1) // limit)

    return render_template(
        "index.html",
        notes=notes,
        query=query,
        page=page,
        max_page=max_page,
        total_count=total_count,
    )


@app.route("/random")
@requires_auth
def random_note():
    """ランダムなノートに飛ばす"""
    conn = get_db_connection()
    try:
        cursor = conn.cursor()
        cursor.execute("SELECT key FROM notes ORDER BY RANDOM() LIMIT 1")
        row = cursor.fetchone()
    finally:
        conn.close()

    if row:
        return redirect(url_for("view_note", key=row["key"]))
    else:
        return redirect(url_for("index"))


@app.route("/tags")
@requires_auth
def tag_list():
    """タグ一覧ページ"""
    # URLパラメータ ?sort=count または ?sort=name を取得 (デフォルトは count)
    sort_mode = request.args.get("sort", "count")

    conn = get_db_connection()
    try:
        # utilsの関数はデフォルトで「件数降順 -> 名前昇順」で返します
        tags = get_all_tags_with_count(conn)
    finally:
        conn.close()

    # 名前順が指定された場合のみ、タグ名(x[0])で再ソート(昇順)
    if sort_mode == "name":
        tags.sort(key=lambda x: x[0])

    return render_template("tags.html", tags=tags, sort_mode=sort_mode)


@app.route("/view/<key>")
@requires_auth
def view_note(key):
    """ノート詳細表示"""
    conn = get_db_connection()
    try:
        notes = fetch_notes_from_db(conn, "key = ?", [key])
    finally:
        conn.close()

    if not notes:
        return "Note not found", 404
    return render_template("view.html", note=notes[0])


@app.route("/pdf/<key>")
@requires_auth
def view_pdf(key):
    """PDFファイルを表示（ダウンロード）"""
    conn = get_db_connection()
    try:
        notes = fetch_notes_from_db(conn, "key = ?", [key])
    finally:
        conn.close()

    if not notes:
        return "Note not found", 404

    note = notes[0]

    # ファイル検索ロジック (Nexus/utils.py のロジックを再現)
    search_paths = []
    if PDF_ROOT:
        search_paths.append(Path(PDF_ROOT))
    if PDF_ARCHIVE:
        search_paths.append(Path(PDF_ARCHIVE))
    if DB_PATH:
        search_paths.append(Path(DB_PATH).parent)

    target_path = None

    # 1. 統合PDFを探す
    merged_name = note.get("merged_pdf_filename")
    if merged_name:
        target_path = find_file_in_paths(merged_name, search_paths)

    # 2. 元ファイルを探す
    if not target_path:
        orig_name = note.get("filepath")
        if orig_name:
            target_path = find_file_in_paths(orig_name, search_paths)

    if target_path and target_path.exists():
        return send_file(target_path)
    else:
        return "PDF File not found on server", 404


# --- サーバー起動ロジック ---
def run_server():
    """サーバー起動エントリーポイント"""

    # 強制終了ハンドラ
    # EXE化したFlask(Windows)はCtrl+Cでゾンビ化しやすいため、
    # シグナルを検知したら os._exit(0) でプロセスごと即座に抹殺する
    def signal_handler(sig, frame):
        print("\nStopping server... (Force Exit)")
        os._exit(0)

    signal.signal(signal.SIGINT, signal_handler)
    signal.signal(signal.SIGTERM, signal_handler)

    print("=" * 50)
    print(" Synapsen Web Server Running...")
    print(" Access URL: http://<This-PC-IP>:5000/")
    print("=" * 50)

    # use_reloader=False は必須 (EXE内でリロード監視が動くとエラーになる/プロセスが増えるため)
    app.run(host="0.0.0.0", port=5000, debug=False, use_reloader=False)


if __name__ == "__main__":
    run_server()

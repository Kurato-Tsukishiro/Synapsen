import sys
import os
import sqlite3
import logging
import signal  # 終了シグナル制御用
import tempfile
import datetime
import re
import shutil
import json
import fitz  # PyMuPDF
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
    flash,  # メッセージ表示用
)

# --- パスの設定 ---
# 親ディレクトリ(Synapsenルート)をパスに追加して Nexus のモジュールを読み込めるようにする
current_dir = Path(__file__).resolve().parent
root_dir = current_dir.parent

if str(root_dir) not in sys.path:
    sys.path.append(str(root_dir))

# Normalisierer のモジュールを読み込めるようにする
normalisierer_dir = root_dir / "Synapsen_Normalisierer"
if str(normalisierer_dir) not in sys.path:
    sys.path.append(str(normalisierer_dir))

# Nexusの既存ロジックをインポート
from Synapsen_Nexus.utils import (  # noqa: E402
    load_app_config,
    fetch_notes_from_db,
    find_file_in_paths,
    count_notes_from_db,
    get_all_tags_with_count,
)
from Synapsen_Nexus.search_parser import parse_query_to_sql  # noqa: E402

# pdf_utils のインポート
try:
    from pdf_utils import (  # type: ignore
        convert_document_to_pdf,
        convert_image_to_pdf,
        high_fidelity_flatten,
        normalize_pdf_to_papersize,
        add_metadata_to_clip,
        embed_processing_flag,
        hex_to_rgb_tuple,
    )
except ImportError as e:
    print(f"Warning: pdf_utils import failed: {e}")


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

app.secret_key = "synapsen_secret_key"

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

    # Watchdogの監視先(Inbox)パスを取得
    import configparser

    if getattr(sys, "frozen", False):
        # .exe実行の場合（config.ini は .exe と同じフォルダ）
        config_path = base_path / "config.ini"
    else:
        # スクリプト実行の場合（config.ini は .py の1つ上のフォルダ）
        config_path = base_path.parent / "config.ini"

    cfg = configparser.ConfigParser(interpolation=None)  # 環境変数の利用を可能にする
    cfg.read(config_path, encoding="utf-8")

    OUTPUT_DIR = cfg.get("Watchdog", "output_dir", fallback=None)

    # デフォルト値を設定
    SERVER_HOST = cfg.get("Server", "host", fallback="0.0.0.0")
    SERVER_PORT = cfg.getint("Server", "port", fallback=5000)  # getintで整数として取得
    USERNAME = cfg.get("Server", "username", fallback="admin")
    PASSWORD = cfg.get("Server", "password", fallback="password")

    # フォントとIndex Keyの取得
    font_path_from_config = cfg.get("Paths", "font_path", fallback="")
    FONT_PATH = os.path.expandvars(font_path_from_config)
    options_str = cfg.get("CommonplaceKeys", "options", fallback="")
    INDEX_KEYS = [k.strip() for k in options_str.split(",") if k.strip()]

    KEY_COLORS = {}
    if cfg.has_section("KeyColors"):
        for k, v in cfg.items("KeyColors"):
            KEY_COLORS[k] = v

    STICKY_COLORS = [
        ("イエロー", "#f8e58c"),
        ("ブルー", "#bbc8e6"),
        ("レッド", "#eebbcb"),
        ("グリーン", "#c1d8ac"),
        ("ホワイト", "#fbfaf5"),
        ("グレー", "#adadad"),
    ]

except Exception as e:
    logger.error(f"設定の読み込みに失敗しました: {e}")
    sys.exit(1)


# --- 保存済み検索の取得ヘルパー ---
def get_saved_searches():
    # exe環境とスクリプト環境の差異を考慮
    saved_searches_path = (
        base_path if getattr(sys, "frozen", False) else base_path.parent
    ) / "saved_searches.json"
    if saved_searches_path.exists():
        try:
            with open(saved_searches_path, "r", encoding="utf-8") as f:
                return json.load(f)
        except Exception as e:
            logger.error(f"保存済み検索の読み込みエラー: {e}")
    return {}


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


# --- アップロード許可設定 ---
ALLOWED_EXTENSIONS = {"pdf", "png", "jpg", "jpeg", "md", "txt"}


def allowed_file(filename):
    return "." in filename and filename.rsplit(".", 1)[1].lower() in ALLOWED_EXTENSIONS


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
    saved_searches = get_saved_searches()  # 保存済み検索を取得

    return render_template(
        "index.html",
        notes=notes,
        query=query,
        page=page,
        max_page=max_page,
        total_count=total_count,
        saved_searches=saved_searches,
    )


@app.route("/create_sticky/<source_key>", methods=["POST"])
@requires_auth
def create_sticky(source_key):
    """付箋ノート(MDファイル)のアップロードとPDF生成"""
    if "file" not in request.files:
        flash("ファイルが選択されていません。", "error")
        return redirect(url_for("view_note", key=source_key))

    file = request.files["file"]
    if file.filename == "":
        flash("ファイルが選択されていません。", "error")
        return redirect(url_for("view_note", key=source_key))

    if not OUTPUT_DIR or not os.path.exists(OUTPUT_DIR):
        flash("正規化済みファイルの出力先(Watchdog)が見つかりません。", "error")
        return redirect(url_for("view_note", key=source_key))

    index_key = request.form.get("index_key", "")
    bg_color = request.form.get("bg_color", "#FFFFA5")

    title = Path(file.filename).stem
    try:
        content = file.read().decode("utf-8")
    except UnicodeDecodeError:
        flash("ファイルの読み込みに失敗しました。UTF-8で保存されていますか？", "error")
        return redirect(url_for("view_note", key=source_key))

    now_str = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
    safe_title = re.sub(r'[\\/:\*\?"<>\|]', "_", title if title else "Sticky")
    base_name = f"{now_str}_{safe_title}"

    try:
        with tempfile.TemporaryDirectory() as temp_dir:
            temp_dir_path = Path(temp_dir)
            md_path = temp_dir_path / f"{base_name}.md"

            # 背景色とマージンを調整するCSS注入
            from pdf_utils import get_sticky_note_css  # type: ignore

            # 共通関数を呼び出してCSSを取得
            style_tag = get_sticky_note_css(bg_color)
            # マークダウンの構成
            md_text = (
                f"{style_tag}\n\n"
                "<div class='content-wrapper'>\n\n"
                f"# {title}\n\n# [内容]\n{content}\n\n</div>"
            )

            with open(md_path, "w", encoding="utf-8") as f:
                f.write(md_text)

            temp_pdf = temp_dir_path / f"temp_{base_name}.pdf"
            temp_flat = temp_dir_path / f"flat_{base_name}.pdf"

            actual_font_path = (
                os.path.expandvars(FONT_PATH)
                if FONT_PATH
                else r"C:\Windows\Fonts\msgothic.ttc"
            )

            # 1. Playwright + Pandoc でPDF化
            convert_document_to_pdf(
                md_path,
                temp_pdf,
                paper_size_str="A4",
                pdf_margins={"top": "0", "bottom": "0", "left": "0", "right": "0"},
            )

            # 2. フラット化
            high_fidelity_flatten(
                str(temp_pdf), str(temp_flat), str(actual_font_path), flatten_ink=False
            )

            final_temp_pdf = temp_dir_path / f"{base_name}.pdf"

            # 3. サイズ正規化
            normalize_pdf_to_papersize(
                str(temp_flat), str(final_temp_pdf), 595.276, 841.89, target_format="A4"
            )

            # 4. 付箋メタデータの書き込み
            try:
                doc = fitz.open(final_temp_pdf)
                meta = doc.metadata
                current_keywords = meta.get("keywords", "")
                new_keywords = (
                    f"{current_keywords}; Synapsen:Sticky"
                    if current_keywords
                    else "Synapsen:Sticky"
                )
                meta["keywords"] = new_keywords
                doc.set_metadata(meta)
                doc.saveIncr()
                doc.close()
            except Exception as e:
                logger.warning(f"付箋識別子の埋め込みに失敗: {e}")

            embed_processing_flag(str(final_temp_pdf))

            # 5. QRコード埋め込み（引用Keyとして元のノートを設定）
            hex_c = KEY_COLORS.get(index_key.lower(), "#000000")
            text_color_rgb = (
                hex_to_rgb_tuple(hex_c) if hex_to_rgb_tuple(hex_c) else (0, 0, 0)
            )

            add_metadata_to_clip(
                pdf_path_str=str(final_temp_pdf),
                font_path=str(actual_font_path),
                paper_width=595.276,
                paper_height=841.89,
                key_rect_tuple=(0, 13, 391, 73),
                index_key_to_embed=index_key,
                text_color=text_color_rgb,
                comment_to_embed=f"Sticky Note: {title}",
                base_name=base_name,
                cited_keys_list=[source_key],  # 元ノートを引用リンクとして設定
                refs_qr_size_pt=75,
                extra_keywords=["Synapsen:Sticky"],
            )

            # 6. 正規化済みファイルフォルダ(Watchdog)へ移動
            shutil.move(str(final_temp_pdf), str(Path(OUTPUT_DIR) / f"{base_name}.pdf"))
            flash(
                f"付箋ノートを作成し、Inboxに送信しました。 ({base_name}.pdf)",
                "success",
            )

    except Exception as e:
        logger.error(f"Sticky creation error: {e}")
        flash(f"付箋作成エラー: {e}", "error")

    return redirect(url_for("view_note", key=source_key))


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


@app.route("/upload", methods=["GET", "POST"])
@requires_auth
def upload():
    """ファイルアップロード画面"""
    if request.method == "POST":
        # ファイルリストを取得
        if "file" not in request.files:
            flash("ファイルが選択されていません。", "error")
            return redirect(request.url)

        # getlist で複数のファイルを取得
        files = request.files.getlist("file")

        # 配列形式でメタデータを取得
        form_index_keys = request.form.getlist("index_keys[]")
        form_comments = request.form.getlist("comments[]")

        if not files or (len(files) == 1 and files[0].filename == ""):
            flash("ファイルが選択されていません。", "error")
            return redirect(request.url)

        if not OUTPUT_DIR or not os.path.exists(OUTPUT_DIR):
            flash(
                f"正規化済みファイルの出力先(Watchdog)が見つかりません設定を確認してください。\nPath: {OUTPUT_DIR}",
                "error",
            )
            return redirect(request.url)

        success_count = 0
        error_count = 0

        actual_font_path = (
            os.path.expandvars(FONT_PATH)
            if FONT_PATH
            else r"C:\Windows\Fonts\msgothic.ttc"
        )

        try:
            with tempfile.TemporaryDirectory() as temp_dir:
                temp_dir_path = Path(temp_dir)

                # enumerate を使用して、ファイルと対応するメタデータを紐付け
                for i, file in enumerate(files):
                    if not (file and allowed_file(file.filename)):
                        error_count += 1
                        continue

                    # 対応するメタデータを取得（安全のため境界チェック）
                    current_index_key = (
                        form_index_keys[i] if i < len(form_index_keys) else ""
                    )
                    current_comment = form_comments[i] if i < len(form_comments) else ""

                    filename = os.path.basename(file.filename)
                    safe_stem = Path(filename).stem
                    ext = Path(filename).suffix.lower()

                    # Synapsenのファイル命名規則に従いタイムスタンプを付与
                    now_str = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
                    base_name = f"{now_str}_{safe_stem}"

                    # オリジナルファイルを一時保存
                    raw_path = temp_dir_path / filename
                    file.save(str(raw_path))

                    temp_pdf = temp_dir_path / f"temp_{base_name}.pdf"
                    temp_flat = temp_dir_path / f"flat_{base_name}.pdf"
                    final_pdf = temp_dir_path / f"{base_name}.pdf"

                    # 1. 変換
                    if ext == ".pdf":
                        shutil.copy2(str(raw_path), str(temp_pdf))
                    elif ext in [".png", ".jpg", ".jpeg"]:
                        convert_image_to_pdf(raw_path, temp_pdf)
                    elif ext in [".md", ".txt"]:
                        convert_document_to_pdf(
                            raw_path,
                            temp_pdf,
                            paper_size_str="A4",
                            pdf_margins={
                                "top": "0",
                                "bottom": "0",
                                "left": "0",
                                "right": "0",
                            },
                        )
                    else:
                        error_count += 1
                        continue

                    # 2. 正規化処理 (フラット化・サイズ調整)
                    high_fidelity_flatten(
                        str(temp_pdf),
                        str(temp_flat),
                        str(actual_font_path),
                        flatten_ink=False,
                    )

                    # 3. サイズ正規化 (余白の調整)
                    normalize_pdf_to_papersize(
                        str(temp_flat),
                        str(final_pdf),
                        595.276,
                        841.89,
                        target_format="A4",
                    )

                    # 二重正規化を防ぐフラグ
                    embed_processing_flag(str(final_pdf))

                    # 4. 個別メタデータの埋め込み
                    hex_c = KEY_COLORS.get(current_index_key.lower(), "#000000")
                    text_color_rgb = (
                        hex_to_rgb_tuple(hex_c)
                        if hex_to_rgb_tuple(hex_c)
                        else (0, 0, 0)
                    )

                    add_metadata_to_clip(
                        pdf_path_str=str(final_pdf),
                        font_path=str(actual_font_path),
                        paper_width=595.276,
                        paper_height=841.89,
                        key_rect_tuple=(0, 13, 391, 73),
                        index_key_to_embed=current_index_key,
                        text_color=text_color_rgb,
                        comment_to_embed=current_comment,
                        base_name=base_name,
                        cited_keys_list=None,
                        refs_qr_size_pt=75,
                    )

                    # 4. 出力フォルダへ移動
                    target_path = Path(OUTPUT_DIR) / f"{base_name}.pdf"
                    shutil.move(str(final_pdf), str(target_path))

                    logger.info(f"Uploaded & Normalized: {filename} -> {target_path}")
                    success_count += 1

            # 結果メッセージの作成
            if success_count > 0:
                msg = f"{success_count} 個のファイルを個別に正規化して保存しました。",
                if error_count > 0:
                    msg += f" (失敗/除外: {error_count} 個)"
                flash(msg, "success")
                return redirect(url_for("index"))
            else:
                flash("処理可能なファイルがありませんでした。", "error")
                return redirect(request.url)

        except Exception as e:
            logger.error(f"Upload Error: {e}")
            flash(f"エラーが発生しました: {e}", "error")
            return redirect(request.url)

    # GET リクエスト時は INDEX_KEYS をテンプレートに渡す
    return render_template("upload.html", index_keys=INDEX_KEYS)


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

    return render_template(
        "view.html", note=notes[0], index_keys=INDEX_KEYS, sticky_colors=STICKY_COLORS
    )


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
    # 設定値を表示に反映
    print(f" Access URL: http://{SERVER_HOST}:{SERVER_PORT}/")
    print("=" * 50)

    # configから読み込んだホストとポートを使用
    app.run(host=SERVER_HOST, port=SERVER_PORT, debug=False, use_reloader=False)


if __name__ == "__main__":
    run_server()

import os
import sys
import customtkinter as ctk
import pandas as pd
import configparser
import webbrowser
import re
from pathlib import Path
from tkinter import messagebox
import sqlite3
import fitz  # PyMuPDF (PDF画像化のために追加)
from PIL import Image  # Pillow (PDF画像化のために追加)
import io  # (PDF画像化のために追加)

import logging
logger = logging.getLogger(__name__)


def load_app_config(base_path):
    """
    config.ini ファイルを読み込み、設定値の辞書を返す。

    Args:
        base_path (Path): アプリケーションの基準パス (main.pyまたは実行ファイルの位置)。

    Returns:
        dict: 読み込まれた設定値の辞書。
               (キー:
                'pdf_root_folder',
                'key_icons',
                'key_colors',

                'commonplace_keys_options',
                'predefined_tags',
                'default_csv_path',
                'include_all_tags_for_autocomplete')

    Raises:
        FileNotFoundError: config.ini が見つからない場合。
        Exception: その他の設定読み込みエラー。
    """
    # .exe実行かスクリプト実行かで config.ini の場所を切り替える
    if getattr(sys, 'frozen', False):
        # .exe実行の場合（config.ini は .exe と同じフォルダ）
        config_path = base_path / 'config.ini'
    else:
        # スクリプト実行の場合（config.ini は .py の1つ上のフォルダ）
        config_path = base_path.parent / 'config.ini'

    config_path = config_path.resolve()  # 絶対パスに正規化
    logger.debug(f"Loading config from: {config_path}")

    if not config_path.is_file():
        raise FileNotFoundError(f"config.iniが見つかりません: {config_path}")

    try:
        parser = configparser.ConfigParser(interpolation=None)
        parser.read(config_path, encoding='utf-8')

        config_data = {}

        # [Paths]
        pdf_root_path_str = parser.get('Paths', 'pdf_root_folder', fallback='')
        if pdf_root_path_str:
            # 環境変数を展開（%LOCALAPPDATA% など）
            pdf_root_path = Path(os.path.expandvars(pdf_root_path_str))
            if not pdf_root_path.is_absolute():
                # config.iniからの相対パスは、config.ini自身からの相対とみなす
                pdf_root_path = config_path.parent / pdf_root_path
            config_data['pdf_root_folder'] = pdf_root_path.resolve()
        else:
            config_data['pdf_root_folder'] = None

        # [KeyIcons]
        if parser.has_section('KeyIcons'):
            config_data['key_icons'] = {
                k.lower(): v for k, v in parser.items('KeyIcons')
            }
        else:
            config_data['key_icons'] = {}

        # [KeyColors]
        if parser.has_section('KeyColors'):
            config_data['key_colors'] = {
                k.lower(): v for k, v in parser.items('KeyColors')
            }
        else:
            config_data['key_colors'] = {}

        # [CommonplaceKeys]
        if parser.has_section('CommonplaceKeys'):
            keys_str = parser.get('CommonplaceKeys', 'options', fallback='')
            config_data['commonplace_keys_options'] = [
                key.strip() for key in keys_str.split(',') if key.strip()
            ]
        else:
            config_data['commonplace_keys_options'] = []

        # [Search] セクションの読み込み
        if parser.has_section('Search'):
            config_data[
                'include_all_tags_for_autocomplete'] = parser.getboolean(
                'Search', 'include_all_tags_for_autocomplete', fallback=True
            )
        else:
            # セクションがない場合のデフォルト値
            config_data['include_all_tags_for_autocomplete'] = True

        # タグリストの読み込み ([Paths] 'tags_data_path')
        config_data['predefined_tags'] = []
        tags_path_from_config = parser.get(
            'Paths', 'tags_data_path', fallback=''
            )
        if tags_path_from_config:
            # 環境変数を展開
            tags_data_path = Path(os.path.expandvars(tags_path_from_config))
            if not tags_data_path.is_absolute():
                # config.iniからの相対パスは、config.ini自身からの相対とみなす
                tags_data_path = config_path.parent / tags_data_path

            try:
                if tags_data_path.is_file():
                    with open(tags_data_path, "r", encoding="utf-8") as f:
                        tags_list = []
                        for line in f:
                            stripped_line = line.strip()

                            # まず空行でないかチェック
                            if stripped_line:
                                # 次にコメント行でないかチェック
                                if not stripped_line.startswith('#'):
                                    tags_list.append(stripped_line)

                        config_data['predefined_tags'] = sorted(tags_list)
            except Exception as e:
                logger.error(f"tags.txtの読み込み中にエラー: {e}")

        # デフォルトDBパスの読み込み ([Paths] 'database_path')
        default_db_path_str = parser.get(
            'Paths', 'database_path', fallback=''
            )
        if default_db_path_str:
            # 環境変数を展開
            db_path = Path(os.path.expandvars(default_db_path_str))
            if not db_path.is_absolute():
                # config.iniからの相対パスは、config.ini自身からの相対とみなす
                db_path = config_path.parent / db_path
            config_data['database_path'] = db_path.resolve()  # <-- キーを変更
        else:
            config_data['database_path'] = None

        return config_data

    except Exception as e:
        # エラーをラップして再度発生させ、呼び出し元 (main.py) で処理する
        raise Exception(f"config.iniの読み込みに失敗しました: {e}")


def load_sql_data_file(filepath: Path):
    """
    指定されたパスからSynapsenのSQLiteデータベースファイルを読み込み、DataFrameを返す。
    【FTS5対応版】 'full_text' と 'memo' を除外してメモリ消費量を削減します。

    Args:
        filepath (Path): 読み込むSQLiteデータベースファイルのパス。

    Returns:
        pd.DataFrame: 読み込まれたデータ。

    Raises:
        Exception: データベースの読み込みまたは処理に失敗した場合。
    """
    if not filepath.is_file():
        # DBファイルが存在しない場合、空のDataFrameを返す
        logger.warning(
            f"データベースファイルが見つかりません: {filepath}",
            extra={'sensitive': True}
        )
        # 空でもカラムは定義しておく
        cols = [
            'tags', 'key', 'memo', 'title', 'commonplace_key', 'date',
            'full_text', 'time', 'pages', 'filepath',
            'merged_pdf_filename', 'merged_start_page'
        ]
        return pd.DataFrame(columns=cols)

    try:
        conn = sqlite3.connect(filepath)
        cursor = conn.cursor()

        # 'notes' テーブルの列情報を取得
        try:
            cursor.execute("PRAGMA table_info(notes)")
            all_columns = [info[1] for info in cursor.fetchall()]
        except sqlite3.OperationalError:
            logger.error(
                f"テーブル 'notes' がDBに存在しません: {filepath}",
                extra={'sensitive': True}
            )
            conn.close()
            cols = ['tags', 'key', 'memo', 'title', 'commonplace_key', 'date',
                    'full_text', 'time', 'pages', 'filepath',
                    'merged_pdf_filename', 'merged_start_page']
            return pd.DataFrame(columns=cols)

        # 'full_text' と 'memo' は除外する
        columns_to_load = [
            col for col in all_columns if col not in ('full_text', 'memo')
        ]

        if not columns_to_load:
            logger.warning(f"テーブル 'notes' に読み込み可能な列がありません: {filepath}")
            conn.close()
            cols = ['tags', 'key', 'memo', 'title', 'commonplace_key', 'date',
                    'full_text', 'time', 'pages', 'filepath',
                    'merged_pdf_filename', 'merged_start_page']
            return pd.DataFrame(columns=cols)

        select_query = f"SELECT {', '.join(columns_to_load)} FROM notes"
        df = pd.read_sql_query(select_query, conn)
        conn.close()

        df = df.fillna('')
        df.columns = df.columns.str.strip()

        # 検索対象となる主要な列を文字列型(str)として明示的に変換
        # 'full_text' 及び 'memo' は型変換リストから除外
        for col in [
            'tags', 'key', 'title', 'commonplace_key', 'date',
            'time', 'pages', 'merged_start_page'
        ]:
            if col in df.columns:
                df[col] = df[col].astype(str)
            else:
                # 読み込まなかったが必須の列を空で追加
                df[col] = ''

        # FTS検索用に 'full_text' と 'memo' を空で定義しておく
        df['full_text'] = ''
        df['memo'] = ''

        return df

    except Exception as e:
        # エラーをラップして呼び出し元 (main.py) で処理する
        raise Exception(f"データベースファイルの読み込みに失敗しました:\n{filepath}\n\n{e}")


def build_memo_display(
        parent_frame,
        memo_text,
        df,
        open_preview_callback,
        frame_width=450
        ):
    """
    メモテキストを解析し、[[key]]リンクをクリック可能なラベルとして
    指定された親フレーム内に動的に構築する。

    Args:
        parent_frame (ctk.CTkFrame or ctk.CTkScrollableFrame):
            ラベルを配置する親ウィジェット。
        memo_text (str): 解析対象のメモテキスト。
        df (pd.DataFrame): ノート全体のDataFrame (リンク先のタイトル検索用)。
        open_preview_callback (callable):
            リンククリック時に呼び出すコールバック関数。
            (例: lambda key: app.open_preview_window(key))
        frame_width (int, optional):
            テキストを折り返すための基準幅。デフォルトは450。
    """
    # 既存のウィジェットをクリア
    for widget in parent_frame.winfo_children():
        widget.destroy()

    pattern = re.compile(r"\[\[(.*?)\]\]")
    last_index = 0

    # ラベルを配置するための内部コンテナ
    # (parent_frame が ScrollableFrame の場合、その中身として機能する)
    content_frame = ctk.CTkFrame(parent_frame, fg_color="transparent")
    content_frame.pack(fill="both", expand=True)

    for match in pattern.finditer(memo_text):
        # 1. リンクより前のテキスト部分
        if non_link_text := memo_text[last_index:match.start()]:
            label = ctk.CTkLabel(
                content_frame, text=non_link_text,
                wraplength=frame_width, justify="left", anchor="w"
                )
            label.pack(fill="x", padx=2, pady=0)

        # 2. リンク部分
        full_match_content = match.group(1).strip()
        # 'key' または 'key: title' の 'key' の部分を取得
        link_key = full_match_content.split(':')[0].strip()

        display_text = f"[[{link_key} (ノート不明)]]"
        if df is not None and not df.empty:
            # key列でリンク先ノートを検索
            linked_note_row = df[df['key'] == link_key]
            if not linked_note_row.empty:
                note_title = linked_note_row.iloc[0].get('title', '（タイトルなし）')
                display_text = f"[[{link_key}: {note_title}]]"

        link_label = ctk.CTkLabel(
            content_frame, text=display_text, text_color="#63B8FF",
            cursor="hand2", wraplength=frame_width, justify="left", anchor="w"
            )
        link_label.pack(fill="x", padx=2, pady=0)
        # リンクにクリックイベントをバインド
        link_label.bind(
            "<Button-1>", lambda e, k=link_key: open_preview_callback(k)
            )

        last_index = match.end()

    # 3. 最後のリンク以降のテキスト部分
    if remaining_text := memo_text[last_index:]:
        label = ctk.CTkLabel(
            content_frame, text=remaining_text,
            wraplength=frame_width, justify="left", anchor="w"
            )
        label.pack(fill="x", padx=2, pady=0)


def _extract_links(memo_text: str) -> set:
    """
    メモテキストから [[key]] または [[key:title]] 形式のリンクを抽出し、
    key のセットを返す。

    Args:
        memo_text (str): 解析対象のメモテキスト。

    Returns:
        set: 抽出されたユニークな 'key' のセット。
    """
    if not memo_text:
        return set()

    link_pattern = re.compile(r"\[\[(.*?)\]\]")
    found_keys = set()

    for match in link_pattern.finditer(memo_text):
        full_match_content = match.group(1).strip()
        # 'key' または 'key: title' の 'key' の部分を取得
        link_key = full_match_content.split(':')[0].strip()
        if link_key:
            found_keys.add(link_key)

    return found_keys


def _update_note_links(
        cursor: sqlite3.Cursor,
        source_key: str,
        new_memo_text: str
):
    """
    note_links テーブルを更新する。(DELETE & INSERT)

    Args:
        cursor (sqlite3.Cursor): トランザクション実行中のカーソル。
        source_key (str): リンク元ノートのKey。
        new_memo_text (str): 新しいメモ本文。
    """
    # 1. 新しいリンク先をメモから解析
    target_keys = _extract_links(new_memo_text)

    # 2. 古いリンクを削除
    cursor.execute(
        "DELETE FROM note_links WHERE source_key = ?",
        (source_key,)
    )

    # 3. 新しいリンクを挿入
    if target_keys:
        links_data = [
            (source_key, target_key) for target_key in target_keys
        ]
        sql_insert_links = (
            "INSERT OR IGNORE INTO note_links (source_key, target_key) "
            "VALUES (?, ?)"
        )
        cursor.executemany(
            sql_insert_links,
            links_data
        )


def build_references_display(
    parent_frame,
    backlinks_df,
    open_preview_callback,
    key_icons,
    key_colors
):
    """
    引用元DataFrameに基づき、クリック可能な引用元リストUIを
    指定された親フレーム内に動的に構築する。

    Args:
        parent_frame (ctk.CTkFrame or ctk.CTkScrollableFrame):
            ラベルを配置する親ウィジェット。
        backlinks_df (pd.DataFrame): 引用元ノートのDataFrame。
        open_preview_callback (callable):
            リンククリック時に呼び出すコールバック関数。
            (例: lambda key: app.open_preview_window(key))
        key_icons (dict): IndexKeyのアイコン辞書。
        key_colors (dict): IndexKeyの色辞書。
    """
    # 既存のウィジェットをクリア
    for widget in parent_frame.winfo_children():
        widget.destroy()

    if backlinks_df.empty:
        parent_frame.configure(label_text="このノートを引用 (0件)")
        return

    parent_frame.configure(
        label_text=f"このノートを引用 ({len(backlinks_df)}件)"
    )

    # (parent_frame が ScrollableFrame の場合、その中身として機能する)
    content_frame = ctk.CTkFrame(parent_frame, fg_color="transparent")
    content_frame.pack(fill="both", expand=True)

    for index, row in backlinks_df.iterrows():
        item_frame = ctk.CTkFrame(content_frame, fg_color="transparent")
        item_frame.pack(fill="x", padx=5, pady=2)

        cp_key = str(row.get("commonplace_key", "")).lower()
        icon = key_icons.get(cp_key, '•')
        color = key_colors.get(cp_key, 'gray')

        icon_label = ctk.CTkLabel(
            item_frame, text=icon, text_color=color,
            font=("", 16), width=20
        )
        icon_label.pack(side="left")

        display_text = f"[{row.get('date')}] {row.get('title', 'N/A')}"
        text_label = ctk.CTkLabel(
            item_frame, text=display_text, anchor="w", cursor="hand2"
        )
        text_label.pack(side="left", fill="x", expand=True)

        # --- Event Binding ---
        key = row.get('key')
        if key:
            def create_preview_handler(note_key=key):
                """ 現在の 'note_key' を保持するイベントハンドラを返す """
                def handler(event):
                    # コールバック関数 (open_preview_callback) を 'note_key' で呼び出す
                    open_preview_callback(note_key)
                return handler

            preview_command = create_preview_handler()
            item_frame.bind("<Button-1>", preview_command)
            icon_label.bind("<Button-1>", preview_command)
            text_label.bind("<Button-1>", preview_command)


def open_pdf_viewer(row_data, loaded_db_path, pdf_root_folder):
    """
    ノートデータに基づき、統合PDFまたは元のPDFを開く。

    Args:
        row_data (pd.Series):
            PDFを開く対象のノートデータ（DataFrameの1行）。
        loaded_db_path (str or Path):  # <-- 'loaded_csv_path' から変更
            現在読み込まれているデータベースのパス (統合PDFの基準パスとして使用)。
        pdf_root_folder (str or Path):
            config.iniで指定された元のPDFのルートフォルダパス。
    """
    merged_pdf_filename = row_data.get('merged_pdf_filename')
    start_page = row_data.get('merged_start_page')

    # 1. 統合PDF (merged_pdf) が指定されている場合
    # pd.isna() で start_page が空 (NaN) でないことも確認
    if merged_pdf_filename and not pd.isna(start_page) and start_page != '':
        if not loaded_db_path:
            messagebox.showerror("エラー", "データベースが読み込まれていないため、PDFの場所を特定できません。")
            return

        # 統合PDFはDBファイルと同じディレクトリにあると想定
        pdf_path = Path(loaded_db_path).parent / merged_pdf_filename

        if not pdf_path.is_file():
            messagebox.showerror("ファイルエラー", f"統合PDFファイルが見つかりません: {pdf_path}")
            return
        try:
            page_number = int(start_page)
            # PDFをページ指定で開くURI
            file_uri = f"{pdf_path.as_uri()}#page={page_number}"
            webbrowser.open(file_uri)
        except (ValueError, TypeError):
            messagebox.showerror("データエラー", "ページ番号が無効です。")
        except Exception as e:
            messagebox.showerror("起動エラー", f"PDFビューアの起動に失敗しました: {e}")

    # 2. 統合PDFがない場合、元のPDF (original_pdf) を試みる
    else:
        _open_original_pdf(row_data, pdf_root_folder)


def get_pdf_uri_for_note(row_data, loaded_db_path, pdf_root_folder):
    """
    ノートデータに基づき、PDFを開くための file:// URI を生成して返す。
    (open_pdf_viewer のロジックを再利用)

    Returns:
        str or None: "file:///path/to/doc.pdf#page=5" 形式のURI。
                     失敗した場合は None。
    """
    merged_pdf_filename = row_data.get('merged_pdf_filename')
    start_page_str = row_data.get('merged_start_page')

    pdf_path = None
    page_number = 1  # 1-indexed

    # 1. 統合PDF (merged_pdf) が指定されている場合
    if (
            merged_pdf_filename and
            not pd.isna(start_page_str) and
            start_page_str != ''
    ):
        if loaded_db_path:
            pdf_path = Path(loaded_db_path).parent / merged_pdf_filename

        if not pdf_path or not pdf_path.is_file():
            pdf_path = None  # 見つからなければ元のPDFを探すフォールバック
        else:
            try:
                page_number = int(start_page_str)
            except (ValueError, TypeError):
                page_number = 1

    # 2. 統合PDFがない (または見つからない) 場合、元のPDF (original_pdf) を試みる
    if pdf_path is None:
        if not pdf_root_folder or not Path(pdf_root_folder).is_dir():
            return None  # 元PDFのルートも不明

        filename = row_data.get('filepath')
        if not filename:
            return None  # 元PDFのファイルパスも不明

        pdf_path = Path(pdf_root_folder) / Path(filename).name
        page_number = 1  # 元PDFは常に1ページ目

        if not pdf_path.is_file():
            return None  # 元PDFが見つからない

    # 3. URIを構築
    try:
        # ページ指定で開くURI
        file_uri = f"{pdf_path.as_uri()}#page={page_number}"
        return file_uri
    except Exception:
        return None


def update_note_in_db(
        conn: sqlite3.Connection,
        key: str,
        new_data: dict
):
    """
    SQLiteデータベース内の指定されたノートを更新し、リンクテーブルも更新する。
    ★注意: この関数は呼び出し元でトランザクション(commit/rollback)を管理する必要がある。

    Args:
        conn (sqlite3.Connection): データベース接続オブジェクト。
        key (str): 更新するノートのユニークID。
        new_data (dict): 更新するデータ ('memo', 'tags', 'commonplace_key')。

    Raises:
        Exception: データベースの更新に失敗した場合。
    """
    # タグをリストから ';' 区切りの文字列に変換
    if 'tags' in new_data and isinstance(new_data['tags'], list):
        tags_str = ";".join(sorted(new_data['tags']))
    else:
        tags_str = new_data.get('tags', '')

    new_memo = new_data.get('memo', '')

    try:
        cursor = conn.cursor()

        # 1. 'notes' テーブル本体を更新
        cursor.execute(
            """
            UPDATE notes
            SET memo = ?, tags = ?, commonplace_key = ?
            WHERE key = ?
            """,
            (
                new_memo,
                tags_str,
                new_data.get('commonplace_key', ''),
                key
            )
        )

        # 2. 'note_links' テーブルを更新
        _update_note_links(cursor, key, new_memo)

    except Exception as e:
        # ロールバックは呼び出し元で行う
        raise Exception(f"DB更新関数(update_note_in_db)でエラー: {e}")


def delete_note_from_db(conn: sqlite3.Connection, key: str):
    """
    SQLiteデータベースから指定されたノートを削除し、関連リンクも削除する。
    ★注意: この関数は呼び出し元でトランザクション(commit/rollback)を管理する必要がある。

    Args:
        conn (sqlite3.Connection): データベース接続オブジェクト。
        key (str): 削除するノートのユニークID。

    Raises:
        Exception: データベースからの削除に失敗した場合。
    """
    try:
        cursor = conn.cursor()

        # 1. 'notes' テーブルから削除 (FTSトリガーが自動でFTS側も削除)
        cursor.execute("DELETE FROM notes WHERE key = ?", (key,))

        # 2. 'note_links' から関連リンクを削除 (リンク元として、リンク先として)
        cursor.execute("DELETE FROM note_links WHERE source_key = ?", (key,))
        cursor.execute("DELETE FROM note_links WHERE target_key = ?", (key,))

    except Exception as e:
        # ロールバックは呼び出し元で行う
        raise Exception(f"DB削除関数(delete_note_from_db)でエラー: {e}")


def _open_original_pdf(row_data, pdf_root_folder):
    """
    元のPDFファイルを開く（open_pdf_viewerの内部ヘルパー）。

    Args:
        row_data (pd.Series): 対象のノートデータ。
        pdf_root_folder (str or Path): 元のPDFのルートフォルダパス。
    """
    if not pdf_root_folder or not Path(pdf_root_folder).is_dir():
        messagebox.showwarning(
            "設定不足",
            "config.iniで指定された 'pdf_root_folder' が見つかりません。" +
            f"\nパス: {pdf_root_folder}"
        )
        return

    filename = row_data.get('filepath')
    if not filename:
        messagebox.showerror("データエラー", "元のファイルパス (filepath列) がデータに含まれていません。")
        return

    # 'filepath' 列のファイル名だけを抽出し、configの'pdf_root_folder'と結合する
    pdf_path = Path(pdf_root_folder) / Path(filename).name

    if not pdf_path.is_file():
        messagebox.showerror("ファイルエラー", f"元のPDFファイルが見つかりません: {pdf_path}")
        return
    try:
        webbrowser.open(pdf_path.as_uri())
    except Exception as e:
        messagebox.showerror("起動エラー", f"PDFビューアの起動に失敗しました: {e}")


def get_pdf_page_image(
    row_data: pd.Series,
    loaded_db_path: Path,
    pdf_root_folder: Path,
    max_width: int = 400
):
    """
    ノートデータに基づき、該当するPDFの1ページ目（または指定ページ）を
    Pillow の Image オブジェクトとして取得し、リサイズして返す。

    Args:
        row_data (pd.Series): 対象のノートデータ。
        loaded_db_path (Path): 読み込まれているDBのパス。
        pdf_root_folder (Path): 元のPDFのルートフォルダパス。
        max_width (int): プレビュー画像の最大幅。

    Returns:
        Image.Image or None: 取得・リサイズされたPillowイメージ。
                             失敗した場合は None。
    """
    pdf_path = None
    page_num_to_open = 0  # 0-indexed

    merged_pdf_filename = row_data.get('merged_pdf_filename')
    start_page_str = row_data.get('merged_start_page')

    # 1. 統合PDF (merged_pdf) が指定されているか
    if merged_pdf_filename and not pd.isna(
            start_page_str) and start_page_str != '':
        if loaded_db_path:
            pdf_path = loaded_db_path.parent / merged_pdf_filename
            try:
                # merged_start_page は 1-indexed なので、0-indexed に変換
                page_num_to_open = int(start_page_str) - 1
            except ValueError:
                page_num_to_open = 0

        if not pdf_path or not pdf_path.is_file():
            logger.error(
                f"[Preview Error] 統合PDFが見つかりません: {pdf_path}",
                extra={'sensitive': True}
            )
            pdf_path = None  # 見つからなければ元のPDFを探すフォールバック

    # 2. 統合PDFがない (または見つからない) 場合、元のPDF (original_pdf) を試みる
    if pdf_path is None:
        if not pdf_root_folder or not pdf_root_folder.is_dir():
            logger.error("[Preview Error] pdf_root_folder が未設定または存在しません。")
            return None

        filename = row_data.get('filepath')
        if not filename:
            logger.error("[Preview Error] filepath がデータに含まれていません。")
            return None

        pdf_path = pdf_root_folder / Path(filename).name
        page_num_to_open = 0  # 元のPDFは常に1ページ目 (index 0)

        if not pdf_path.is_file():
            logger.error(
                f"[Preview Error] 元のPDFファイルが見つかりません: {pdf_path}",
                extra={'sensitive': True}
            )
            return None

    # 3. PyMuPDFでPDFを開き、ページを画像化
    doc = None
    try:
        doc = fitz.open(pdf_path)
        if not (0 <= page_num_to_open < len(doc)):
            logger.error(f"[Preview Error] 無効なページ番号: {page_num_to_open}")
            if doc:
                doc.close()
            return None

        page = doc[page_num_to_open]

        # ページをPixmap（ピクセルマップ）としてレンダリング (DPI=150)
        # DPIを上げると高画質になるが遅くなる
        pix = page.get_pixmap(dpi=150)

        # Pixmapからバイトデータを取得
        img_data = pix.tobytes("png")
        if not img_data:
            if doc:
                doc.close()
            return None

        # バイトデータからPillowイメージを開く
        pil_image = Image.open(io.BytesIO(img_data))

        # 4. アスペクト比を維持してリサイズ
        if pil_image.width > max_width:
            scale = max_width / pil_image.width
            new_height = int(pil_image.height * scale)
            pil_image = pil_image.resize(
                (max_width, new_height), Image.Resampling.LANCZOS
            )

        return pil_image

    except Exception as e:
        logger.error(
            f"[Preview Error] PDFの画像化に失敗 ({pdf_path}): {e}",
            extra={'sensitive': True}
        )
        return None
    finally:
        if doc:
            doc.close()


def get_pdf_document_for_note(
    row_data: pd.Series,
    loaded_db_path: Path,
    pdf_root_folder: Path
) -> tuple[fitz.Document | None, int, int]:
    """
    ノートデータに紐づくPDFドキュメント(fitz)と、
    表示すべき開始ページインデックス、総ページ数を返す。

    Returns:
        tuple[fitz.Document | None, int, int]:
            (doc, start_page_index, total_pages)
            doc: PyMuPDFのドキュメントオブジェクト (失敗時はNone)
            start_page_index: 表示すべき開始ページ (0-indexed)
            total_pages: このノートの総ページ数 (統合PDFの場合は pages 列の値)
    """
    pdf_path = None
    start_page_index = 0  # 0-indexed
    total_pages = 0

    merged_pdf_filename = row_data.get('merged_pdf_filename')
    start_page_str = row_data.get('merged_start_page')

    # 1. 統合PDF (merged_pdf) が指定されているか
    if (
        merged_pdf_filename and
        loaded_db_path and
        not pd.isna(start_page_str) and
        start_page_str != ''
    ):
        pdf_path = loaded_db_path.parent / merged_pdf_filename
        if not pdf_path.is_file():
            pdf_path = None  # 見つからなければ元のPDFを探す
        else:
            try:
                # 1-indexed -> 0-indexed
                start_page_index = int(start_page_str) - 1
            except ValueError:
                start_page_index = 0

    # 2. 統合PDFがない (または見つからない) 場合、元のPDF (original_pdf) を試みる
    if pdf_path is None:
        if not pdf_root_folder or not pdf_root_folder.is_dir():
            logger.debug("[utils] pdf_root_folder が未設定または存在しません。")
            return (None, 0, 0)
        filename = row_data.get('filepath')
        if not filename:
            logger.debug("[utils] filepath がデータに含まれていません。")
            return (None, 0, 0)

        pdf_path = pdf_root_folder / Path(filename).name
        start_page_index = 0  # 元のPDFは常に 0 から
        if not pdf_path.is_file():
            logger.debug(
                f"[utils] 元のPDFファイルが見つかりません: {pdf_path}",
                extra={'sensitive': True}
            )
            return (None, 0, 0)

    # 3. PyMuPDFでPDFを開く
    try:
        doc = fitz.open(pdf_path)
        doc_total_pages = len(doc)

        # 開始ページがドキュメントの総ページ数を超えていたら0に戻す
        if not (0 <= start_page_index < doc_total_pages):
            start_page_index = 0

        # このノートが担当するページ数 (pages列) を取得
        note_page_count_str = str(row_data.get('pages', '0'))
        if note_page_count_str.isdigit():
            note_page_count = int(note_page_count_str)
        else:
            note_page_count = 0

        # 統合PDFの場合、total_pages を ノートのページ数(pages列) に設定
        # (例: 統合PDF全体は300ページでも、このノートは3ページ分)
        if merged_pdf_filename and note_page_count > 0:
            # ただし、ドキュメントの物理ページ数を超えないようにする
            available_pages = doc_total_pages - start_page_index
            total_pages = min(note_page_count, available_pages)
        else:
            # 元PDFの場合、ドキュメント全体のページ数をそのまま使用
            total_pages = doc_total_pages

        return (doc, start_page_index, total_pages)

    except Exception as e:
        logger.error(
            f"[utils] PDFドキュメントの読み込みに失敗 ({pdf_path}): {e}",
            extra={'sensitive': True}
        )
        return (None, 0, 0)


def get_pdf_page_image_from_doc(
    doc: fitz.Document,
    page_index: int,
    max_width: int = 400
) -> Image.Image | None:
    """
    既に開かれているfitz.Documentオブジェクトとページインデックスから
    Pillowイメージを取得する。
    (旧 get_pdf_page_image からロジックを移植・変更)

    Args:
        doc (fitz.Document): PyMuPDFのドキュメントオブジェクト。
        page_index (int): 描画するページのインデックス (0-indexed)。
        max_width (int): プレビュー画像の最大幅。

    Returns:
        Image.Image or None: 取得・リサイズされたPillowイメージ。
    """
    try:
        if not (0 <= page_index < len(doc)):
            logger.error(f"[utils] 無効なページインデックス: {page_index}")
            return None

        page = doc[page_index]
        # DPIを150に設定 (プレビュー品質)
        pix = page.get_pixmap(dpi=150)
        img_data = pix.tobytes("png")
        if not img_data:
            return None

        pil_image = Image.open(io.BytesIO(img_data))

        # アスペクト比を維持してリサイズ
        scale = max_width / pil_image.width
        new_height = int(pil_image.height * scale)
        pil_image = pil_image.resize(
            (max_width, new_height), Image.Resampling.LANCZOS
        )

        return pil_image

    except Exception as e:
        logger.error(f"[utils] PDFの画像化に失敗 (Page {page_index}): {e}")
        return None

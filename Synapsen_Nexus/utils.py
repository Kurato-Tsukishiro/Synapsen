import customtkinter as ctk
import pandas as pd
import configparser
import webbrowser
import re
from pathlib import Path
from tkinter import messagebox


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
                'default_csv_path')

    Raises:
        FileNotFoundError: config.ini が見つからない場合。
        Exception: その他の設定読み込みエラー。
    """
    # config.ini は 'src' フォルダの一つ上にあると想定
    config_path = base_path.parent / 'config.ini'
    config_path = config_path.resolve()  # 絶対パスに正規化

    if not config_path.is_file():
        raise FileNotFoundError(f"config.iniが見つかりません: {config_path}")

    try:
        parser = configparser.ConfigParser()
        parser.read(config_path, encoding='utf-8')

        config_data = {}

        # [Paths]
        pdf_root_path_str = parser.get('Paths', 'pdf_root_folder', fallback='')
        if pdf_root_path_str:
            pdf_root_path = Path(pdf_root_path_str)
            if not pdf_root_path.is_absolute():
                # config.iniからの相対パスは、config.ini自身からの相対とみなす
                pdf_root_path = config_path.parent / pdf_root_path
            config_data['pdf_root_folder'] = pdf_root_path.resolve()
        else:
            config_data['pdf_root_folder'] = None  # もし空ならNoneにする

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

        # タグリストの読み込み ([Paths] 'tags_data_path')
        config_data['predefined_tags'] = []
        tags_path_from_config = parser.get(
            'Paths', 'tags_data_path', fallback=''
            )
        if tags_path_from_config:
            tags_data_path = Path(tags_path_from_config)
            if not tags_data_path.is_absolute():
                # config.iniからの相対パスは、config.ini自身からの相対とみなす
                tags_data_path = config_path.parent / tags_data_path

            try:
                if tags_data_path.is_file():
                    with open(tags_data_path, "r", encoding="utf-8") as f:
                        config_data['predefined_tags'] = \
                            sorted([line.strip() for line in f if line.strip() and not line.startswith('#')])
            except Exception as e:
                print(f"tags.txtの読み込み中にエラー: {e}")

        # デフォルトCSVパスの読み込み ([Paths] 'default_csv_path')
        default_csv_path_str = parser.get(
            'Paths', 'default_csv_path', fallback=''
            )
        if default_csv_path_str:
            csv_path = Path(default_csv_path_str)
            if not csv_path.is_absolute():
                # config.iniからの相対パスは、config.ini自身からの相対とみなす
                csv_path = config_path.parent / csv_path
            config_data['default_csv_path'] = csv_path.resolve()
        else:
            config_data['default_csv_path'] = None

        return config_data

    except Exception as e:
        # エラーをラップして再度発生させ、呼び出し元 (main.py) で処理する
        raise Exception(f"config.iniの読み込みに失敗しました: {e}")


def load_csv_data_file(filepath):
    """
    指定されたパスから目次CSVファイルを読み込み、DataFrameを返す。
    必須列は文字列型(str)に変換する。

    Args:
        filepath (str or Path): 読み込むCSVファイルのパス。

    Returns:
        pd.DataFrame: 読み込まれたデータ。

    Raises:
        Exception: CSVファイルの読み込みまたは処理に失敗した場合。
    """
    try:
        df = pd.read_csv(filepath, encoding='utf-8-sig').fillna('')
        df.columns = df.columns.str.strip()

        # 検索対象となる主要な列を文字列型(str)として明示的に変換
        # これにより、数値キーなどが検索できなくなる問題を回避する
        for col in ['tags', 'key', 'memo', 'title', 'commonplace_key', 'date']:
            if col in df.columns:
                df[col] = df[col].astype(str)
            else:
                # 必須列がない場合は空の列を追加
                df[col] = ''

        return df
    except Exception as e:
        # エラーをラップして呼び出し元 (main.py) で処理する
        raise Exception(f"CSVファイルの読み込みに失敗しました:\n{filepath}\n\n{e}")


def build_memo_display(parent_frame, memo_text, df, open_preview_callback, frame_width=450):
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


def open_pdf_viewer(row_data, loaded_csv_path, pdf_root_folder):
    """
    ノートデータに基づき、統合PDFまたは元のPDFを開く。

    Args:
        row_data (pd.Series):
            PDFを開く対象のノートデータ（DataFrameの1行）。
        loaded_csv_path (str or Path):
            現在読み込まれている目次CSVのパス (統合PDFの基準パスとして使用)。
        pdf_root_folder (str or Path):
            config.iniで指定された元のPDFのルートフォルダパス。
    """
    merged_pdf_filename = row_data.get('merged_pdf_filename')
    start_page = row_data.get('merged_start_page')

    # 1. 統合PDF (merged_pdf) が指定されている場合
    # pd.isna() で start_page が空 (NaN) でないことも確認
    if merged_pdf_filename and not pd.isna(start_page) and start_page != '':
        if not loaded_csv_path:
            messagebox.showerror("エラー", "CSVファイルが読み込まれていないため、PDFの場所を特定できません。")
            return

        # 統合PDFはCSVファイルと同じディレクトリにあると想定
        pdf_path = Path(loaded_csv_path).parent / merged_pdf_filename

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

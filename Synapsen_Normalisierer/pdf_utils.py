import fitz  # PyMuPDF
from pypdf import PdfReader, PdfWriter, Transformation
from pathlib import Path
import pytesseract
from PIL import Image
import io
import pandas as pd
from pytesseract import Output
import csv
import shutil
import sys
import re
import subprocess

from playwright.sync_api import sync_playwright, Error as PlaywrightError

# ==============================================================================
# 定数定義 (ポイント単位)
# ==============================================================================

CM_TO_PT: float = 72 / 2.54
MARGIN_CM: float = 0
HEAD_SEP_CM: float = 1.0
MARGIN: float = MARGIN_CM * CM_TO_PT
HEAD_SEP: float = HEAD_SEP_CM * CM_TO_PT
TOP_MARGIN: float = MARGIN + HEAD_SEP
BOTTOM_MARGIN: float = MARGIN
LEFT_MARGIN: float = 0
RIGHT_MARGIN: float = 0
# ==============================================================================


def hex_to_rgb_tuple(hex_color: str) -> tuple[float, float, float] | None:
    """
    #RRGGBB 形式の16進数カラーコードを、
    fitzが要求する (R, G, B) のタプル (各値 0.0～1.0) に変換します。
    (webclip_window.py から移植)

    Args:
        hex_color (str): 16進数カラーコード (例: "#FF0000")。

    Returns:
        tuple[float, float, float] | None:
            fitz用のRGBタプル。変換失敗時はNone。
    """
    try:
        hex_color = hex_color.lstrip('#')
        # 16進数を 0-255 の整数に変換
        r, g, b = tuple(int(hex_color[i:i+2], 16) for i in (0, 2, 4))
        # 0-1 の浮動小数点数に変換
        return (r / 255.0, g / 255.0, b / 255.0)
    except Exception as e:
        print(f"警告: 16進数カラーコード '{hex_color}' の変換に失敗: {e}")
        return None


def add_metadata_to_web_clip(
    pdf_path_str: str,
    font_path: str,
    paper_width: float,
    paper_height: float,
    key_rect_tuple: tuple,
    index_key_to_embed: str,
    text_color: tuple | None,
    comment_to_embed: str,
    sist_string_formal: str | None = None,
    sist_string_readable: str | None = None
) -> None:
    """
    [Webクリップ専用]
    Playwrightで生成されたPDFに対し、
    1ページ目に IndexKey を、最終ページ（新規追加）に コメントと書誌情報 を書き込みます。

    Args:
        pdf_path_str (str): 処理対象のPDFファイルパス (読み書きされる)。
        font_path (str): 埋め込むフォントファイルのパス。
        paper_width (float): 最終ページの幅 (ポイント)。
        paper_height (float): 最終ページの高さ (ポイント)。
        key_rect_tuple (tuple): IndexKeyを描画する座標 (x0, y0, x1, y1)。
        index_key_to_embed (str): 描画するIndexKeyの文字列。
        text_color (tuple | None): IndexKeyの文字色 (fitz形式)。
        comment_to_embed (str): 描画するコメント文字列。
        sist_string_formal (str | None): 書誌情報(SIST)。
        sist_string_readable (str | None): 書誌情報(可読形式)。

    Raises:
        Exception: PDFの読み書きやフォント埋め込みに失敗した場合。
    """

    # --- 埋め込む情報が何もなければ、処理をスキップ ---
    if (not index_key_to_embed and
        not comment_to_embed and
            not sist_string_formal):
        print(f"  [Info] 埋め込むメタデータがないためスキップ: {Path(pdf_path_str).name}")
        return

    doc = None
    try:
        doc = fitz.open(pdf_path_str)
        if len(doc) == 0:
            print(f"  [Error] メタデータ埋め込みスキップ: ページが存在しません {pdf_path_str}")
            return

        key_rect = fitz.Rect(key_rect_tuple)
        font_alias = "embed_font"

        # --- 1. 1ページ目に IndexKey を描画 ---
        if index_key_to_embed:
            page1 = doc[0]  # 既存の1ページ目（Webページ本体）を取得
            try:
                page1.insert_font(fontname=font_alias, fontfile=font_path)
            except Exception as e:
                print(f"フォント埋め込み警告 (Page 1): {e}")

            shape1 = page1.new_shape()
            shape1.insert_textbox(
                key_rect, index_key_to_embed, fontname=font_alias,
                fontsize=10, color=text_color, align=0
            )
            shape1.commit()

        # --- 2. 最終ページに コメントと書誌情報 を描画 ---
        if comment_to_embed or sist_string_formal or sist_string_readable:

            # 既存のページ数 (len(doc)) をpnoに指定し、末尾に新しい白紙ページを挿入
            page_last = doc.new_page(
                pno=len(doc), width=paper_width, height=paper_height
            )
            try:
                page_last.insert_font(fontname=font_alias, fontfile=font_path)
            except Exception as e:
                print(f"フォント埋め込み警告 (Last Page): {e}")

            shape_last = page_last.new_shape()

            # --- 描画座標の計算 (embed_metadata_as_cover_page と同様) ---
            info_rect_y_start = key_rect.y1 + 10
            info_rect_y_end = page_last.rect.height - 30
            current_y_pos = info_rect_y_start
            x0 = key_rect.x0

            # 2a. 書誌情報 (SIST 02 単一行)
            if sist_string_formal:
                sist_rect = fitz.Rect(
                    x0, current_y_pos,
                    page_last.rect.width - 50, current_y_pos + 60
                )
                rc_sist = shape_last.insert_textbox(
                    sist_rect, f"書誌情報 (SIST 02):\n{sist_string_formal}",
                    fontname=font_alias, fontsize=6, align=0
                )
                actual_sist_y1 = (
                    sist_rect.y0 + (sist_rect.height - rc_sist)
                    if rc_sist >= 0 else sist_rect.y1
                )
                current_y_pos = actual_sist_y1 + 10  # Y座標を更新

            # 2b. 書誌情報 (改行形式)
            if sist_string_readable:
                readable_rect = fitz.Rect(
                    x0, current_y_pos,
                    page_last.rect.width - 50, info_rect_y_end
                )
                rc_readable = shape_last.insert_textbox(
                    readable_rect, f"書誌情報:\n{sist_string_readable}",
                    fontname=font_alias, fontsize=9, align=0
                )
                actual_readable_y1 = (
                    readable_rect.y0 + (readable_rect.height - rc_readable)
                    if rc_readable >= 0 else readable_rect.y1
                )
                current_y_pos = actual_readable_y1 + 40  # Y座標を更新
            else:
                current_y_pos = info_rect_y_start + 40

            # 2c. コメント (D&D / WebClip 共通)
            comment_rect = fitz.Rect(
                x0, current_y_pos,
                page_last.rect.width - 50, info_rect_y_end
            )
            if comment_to_embed:
                shape_last.insert_textbox(
                    comment_rect, f"コメント:\n{comment_to_embed}",
                    fontname=font_alias, fontsize=9, align=0
                )

            # 描画内容をコミット
            shape_last.commit()

        # 変更を上書き保存
        doc.saveIncr()

    except Exception as e:
        print(f"  [Error] Webクリップへのメタデータ埋め込み中にエラー ({pdf_path_str}): {e}")
        raise  # エラーを再送出
    finally:
        if doc:
            doc.close()


def add_metadata_to_image_clip(
    pdf_path_str: str,
    font_path: str,
    paper_width: float,
    paper_height: float,
    key_rect_tuple: tuple,
    index_key_to_embed: str,
    text_color: tuple | None,
    comment_to_embed: str
) -> None:
    """
    [新設 / D&D・画像クリップ用]
    正規化済みの画像PDF（1ページ目=画像）に対し、
    1ページ目に IndexKey を、
    2ページ目（新規追加）に コメント を書き込みます。

    Args:
        pdf_path_str (str): 処理対象のPDFファイルパス (読み書きされる)。
        font_path (str): 埋め込むフォントファイルのパス。
        paper_width (float): 表紙ページの幅 (ポイント)。
        paper_height (float): 表紙ページの高さ (ポイント)。
        key_rect_tuple (tuple): IndexKeyを描画する座標 (x0, y0, x1, y1)。
        index_key_to_embed (str): 描画するIndexKeyの文字列。
        text_color (tuple | None): IndexKeyの文字色 (fitz形式)。
        comment_to_embed (str): 描画するコメント文字列。

    Raises:
        Exception: PDFの読み書きやフォント埋め込みに失敗した場合。
    """

    # --- 埋め込む情報が何もなければ、処理をスキップ ---
    if not index_key_to_embed and not comment_to_embed:
        print(f"  [Info] 埋め込むメタデータがないためスキップ: {Path(pdf_path_str).name}")
        return

    doc = None
    try:
        doc = fitz.open(pdf_path_str)
        if len(doc) == 0:
            print(f"  [Error] メタデータ埋め込みスキップ: ページが存在しません {pdf_path_str}")
            return

        key_rect = fitz.Rect(key_rect_tuple)
        font_alias = "embed_font"

        # --- 1. 1ページ目に IndexKey を描画 ---
        if index_key_to_embed:
            page1 = doc[0]  # 既存の1ページ目（画像）を取得
            try:
                page1.insert_font(fontname=font_alias, fontfile=font_path)
            except Exception as e:
                print(f"フォント埋め込み警告 (Page 1): {e}")

            # 1ページ目に直接 IndexKey を描画
            shape1 = page1.new_shape()
            shape1.insert_textbox(
                key_rect, index_key_to_embed, fontname=font_alias,
                fontsize=10, color=text_color, align=0
            )
            shape1.commit()

        # --- 2. 2ページ目に コメント を描画 ---
        if comment_to_embed:
            # 2ページ目 (pno=1) に新しい白紙ページを挿入
            page2 = doc.new_page(pno=1, width=paper_width, height=paper_height)
            try:
                page2.insert_font(fontname=font_alias, fontfile=font_path)
            except Exception as e:
                print(f"フォント埋め込み警告 (Page 2): {e}")

            # 描画座標 (IndexKeyと同じマージン)
            comment_rect = fitz.Rect(
                key_rect.x0,
                key_rect.y1 + 10,  # IndexKey描画位置の下から
                page2.rect.width - 50,
                page2.rect.height - 30
            )

            shape2 = page2.new_shape()
            shape2.insert_textbox(
                comment_rect, f"コメント:\n{comment_to_embed}",
                fontname=font_alias, fontsize=9, align=0
            )
            shape2.commit()

        # 変更を上書き保存
        doc.saveIncr()

    except Exception as e:
        print(f"  [Error] 画像クリップへのメタデータ埋め込み中にエラー ({pdf_path_str}): {e}")
        raise  # エラーを再送出
    finally:
        if doc:
            doc.close()


def high_fidelity_flatten(
        input_path: str,
        output_path: str,
        font_path: str
) -> None:
    """
    PyMuPDFを使い、指定フォントでフォームをテキストに変換（高精度フラット化）します。

    Acrobatの「フォームをフラット化」とは異なり、
    フォームフィールドの「値」を指定フォントでベタ書きし、
    フィールド自体を削除することで、注釈（アノテーション）を維持します。

    Args:
        input_path (str): 入力PDFファイルのパス。
        output_path (str): フラット化後の出力PDFファイルのパス。
        font_path (str): 埋め込むフォントファイル（.ttf, .otfなど）のパス。

    Raises:
        FileNotFoundError: 指定されたフォントファイルが見つからない場合。
        Exception: PDFのオープンや保存に失敗した場合。
    """
    if not Path(font_path).is_file():
        raise FileNotFoundError(f"指定されたフォントファイルが見つかりません: {font_path}")

    doc = None
    try:
        doc = fitz.open(input_path)
        if doc.is_encrypted:
            print(f"  [Warn]暗号化されたPDFはスキップします: {Path(input_path).name}")
            return  # 暗号化ファイルは処理せず終了

        font_name_in_pdf = "synapsen-embed-font"  # PDF内部で使うフォントのエイリアス名

        for page in doc:
            try:
                page.insert_font(fontname=font_name_in_pdf, fontfile=font_path)
            except Exception as e:
                print(f"Info: Font insertion issue ({e}). Continuing.")

            # フォームウィジェットを処理
            for widget in page.widgets():
                if widget.field_type in (
                    fitz.PDF_WIDGET_TYPE_TEXT,
                    fitz.PDF_WIDGET_TYPE_COMBOBOX
                ) and widget.field_value:
                    # フィールドの値をページに直接描画
                    page.insert_textbox(
                        widget.rect,
                        widget.field_value,
                        fontname=font_name_in_pdf,
                        fontsize=widget.text_fontsize or 10,
                        color=widget.text_color or (0, 0, 0),
                    )
                # 元のインタラクティブなウィジェットを削除
                page.delete_widget(widget)

        # PDFを保存 (ガベージコレクション、圧縮を有効化)
        doc.save(output_path, garbage=4, deflate=True)

    except Exception as e:
        print(f"  [Error] フラット化処理中にエラー ({input_path}): {e}")
        # エラーが発生した場合も、finally で doc.close() が呼ばれる
        raise  # エラーを再送出

    finally:
        if doc:
            doc.close()


def normalize_pdf_to_papersize(
        input_path: str,
        output_path: str,
        paper_width: float,
        paper_height: float
        ) -> None:
    """
    pypdfを使い、PDFの全ページを、指定された用紙サイズの中央にリサイズ・配置します。

    マージン領域 (TOP_MARGIN, BOTTOM_MARGIN など) を考慮し、
    コンテンツがその領域内に収まるようにアスペクト比を維持して
    スケーリングおよび中央配置を行います。

    Args:
        input_path (str): 入力PDFファイル（通常はフラット化済み）のパス。
        output_path (str): 正規化された出力PDFファイルのパス。
        paper_width (float): ターゲットの用紙幅 (ポイント単位)。
        paper_height (float): ターゲットの用紙高 (ポイント単位)。

    Raises:
        Exception: PDFの読み込み、書き込みに失敗した場合。
    """
    reader = None
    try:
        reader = PdfReader(input_path)
        writer = PdfWriter()

        # 渡された用紙サイズから描画可能領域を計算
        drawable_width: float = paper_width - LEFT_MARGIN - RIGHT_MARGIN
        drawable_height: float = paper_height - TOP_MARGIN - BOTTOM_MARGIN

        for content_page in reader.pages:
            # 指定された用紙サイズの白紙ページを作成
            template_page = writer.add_blank_page(
                width=paper_width, height=paper_height)

            original_width = float(content_page.mediabox.width)
            original_height = float(content_page.mediabox.height)

            if original_width == 0 or original_height == 0:
                print(f"Skipping empty or invalid page in {input_path}")
                continue

            # 描画可能領域 (drawable_width, drawable_height) に収まるようスケーリング
            scale = min(
                drawable_width / original_width,
                drawable_height / original_height
            )

            # 描画可能領域内で中央に配置
            tx = LEFT_MARGIN + (drawable_width - original_width * scale) / 2
            ty = (
                BOTTOM_MARGIN + (drawable_height - original_height * scale) / 2
            )

            transform = (
                Transformation()
                .scale(sx=scale, sy=scale)
                .translate(tx=tx, ty=ty)
            )

            template_page.merge_transformed_page(content_page, transform)

        with open(output_path, "wb") as f:
            writer.write(f)

    except Exception as e:
        print(f"  [Error] 正規化処理中にエラー ({input_path}): {e}")
        raise

    # pypdf (PdfReader) は明示的な close() を必要としない


def embed_ocr_text_in_pdf(
    pdf_path_str: str,
    enable_tesseract: bool,
    font_path: str,
    lang: str = 'jpn+jpn_vert'
) -> None:
    """
    PDFを解析し、既存のテキストレイヤーが存在しない場合、
    かつ `enable_tesseract` が True の場合にのみ Tesseract OCRを実行し、
    結果を「透明なテキストレイヤー」としてPDF自体（指定されたパス）に上書き保存します。

    Args:
        pdf_path_str (str): 処理対象のPDFファイルパス（読み書きされる）。
        enable_tesseract (bool): Tesseract OCR (低速) を実行するかどうか。
        font_path (str): 埋め込む日本語フォントファイルのパス。
        lang (str): Tesseractが使用する言語。

    Raises:
        Exception: Tesseract-OCRが見つからない場合。
                   その他、ファイルの上書き保存に失敗した場合。
    """
    doc = None
    # 一時ファイルへの保存パスを定義
    temp_output_path = pdf_path_str + "._temp_ocr.pdf"
    OCR_FONT_NAME = "synapsen_ocr_font"  # 埋め込みフォントのエイリアス

    try:
        doc = fitz.open(pdf_path_str)
        if doc.is_encrypted:
            print(f"  [Info] 暗号化されたPDFはスキップします: {Path(pdf_path_str).name}")
            return

        # 1. 高速なテキスト抽出を試みる
        meaningful_text_threshold = 10

        if not enable_tesseract:
            print("  [Info] Tesseract OCR は無効です。スキップします。")
            return

        print(
            "  [Info] Tesseract OCR を実行し、テキストを埋め込みます: "
            f"{Path(pdf_path_str).name}"
        )

        pages_processed_count = 0

        for page_num, page in enumerate(doc):

            # 1. ページごとに既存テキストをチェック
            page_text = page.get_text("text", sort=True).strip()
            if len(page_text) > meaningful_text_threshold:
                print(f"  [Info] Page {page_num + 1} には既存テキストがあるためスキップ。")
                continue

            # --- 既存テキストがないページのみ、以下を実行 ---
            print(f"  [Info] Tesseract OCR を実行中 (Page {page_num + 1})...")
            pages_processed_count += 1

            try:
                # 6. 高解像度 (DPI=300) でページを画像(Pixmap)にレンダリング
                pix = page.get_pixmap(dpi=300)
                img_data = pix.tobytes("png")
                img = Image.open(io.BytesIO(img_data))

                # 7. Tesseract OCR の実行 (TSVデータとして取得)
                tsv_data = pytesseract.image_to_data(
                    img, lang=lang, output_type=Output.STRING
                )

                if tsv_data is None or len(tsv_data.strip()) == 0:
                    print(
                        "  [Warn] TesseractがTSVデータを返しませんでした "
                        f"(Page {page_num + 1})"
                    )
                    continue

                # 8. TesseractのTSVデータを解析
                df = pd.read_csv(
                    io.StringIO(tsv_data),
                    sep='\t',
                    quoting=csv.QUOTE_NONE,
                    on_bad_lines='skip'
                )
                df = df.dropna(subset=['conf', 'text'])
                df = df[df['conf'] > 30]  # 信頼度が低いものは除外

                if df.empty:
                    print(
                        "  [Info] Tesseract OCR は実行されましたが、"
                        "埋め込み可能なテキスト(conf > 30)が見つかりませんでした "
                        f"(Page {page_num + 1})。")
                    continue

                # 9. ページに日本語フォントを登録
                try:
                    page.insert_font(
                        fontname=OCR_FONT_NAME, fontfile=font_path)
                except Exception:
                    pass  # 既に登録済みなどのエラーは無視

                # 10. 透明テキストを挿入
                dpi_scale = 72 / 300  # DPI=300 -> 72 DPI (ポイント) に座標を戻す
                for _, row in df.iterrows():
                    x0, y0, w, h = (
                        row['left'], row['top'], row['width'], row['height']
                    )
                    rect = fitz.Rect(
                        x0 * dpi_scale, y0 * dpi_scale,
                        (x0 + w) * dpi_scale, (y0 + h) * dpi_scale
                    )
                    fs = max(h * dpi_scale * 0.8, 6.0)  # フォントサイズ

                    page.insert_text(
                        rect.bottom_left,
                        str(row['text']),
                        fontname=OCR_FONT_NAME,
                        fontsize=fs,
                        render_mode=3,  # 3 = 透明 (描画せず、テキスト選択・検索のみ可能)
                        rotate=0
                    )

            except pytesseract.TesseractNotFoundError:
                # このエラーは回復不能なので、ループを抜けて上位に投げる
                raise Exception("Tesseract-OCRが見つかりません。PATHを確認してください。")
            except Exception as ocr_err:
                print(
                    "  [Warn] Tesseract OCRエラー (Page {page_num + 1}): "
                    f"{ocr_err}")
                continue

        if pages_processed_count == 0:
            print("  [Info] OCRが実行されたページはありませんでした。")
            return

        # 11. 変更を「一時ファイル」に保存
        doc.save(
            temp_output_path,
            garbage=4,
            deflate=True,
            encryption=fitz.PDF_ENCRYPT_NONE
        )
        print(f"  [Info] テキスト埋め込み完了 (一時ファイル): {Path(temp_output_path).name}")

    except Exception as e:
        print(f"  [Error] PDFテキスト埋め込み処理中にエラー ({pdf_path_str}): {e}")
        if Path(temp_output_path).is_file():
            try:
                Path(temp_output_path).unlink()  # エラー時は一時ファイルを削除
            except Exception as e_del:
                print(f"  [Warn] エラー発生後の一時ファイル削除に失敗: {e_del}")
        if doc:
            doc.close()
        raise  # TesseractNotFoundError などを上位に伝える

    finally:
        if doc:
            doc.close()

    # 12. 正常終了した場合のみ、ファイルのリネーム（上書き）
    if Path(temp_output_path).is_file():
        try:
            shutil.move(temp_output_path, pdf_path_str)
            print(f"  [Info] 元ファイルに上書き完了: {Path(pdf_path_str).name}")
        except Exception as e_move:
            print(f"  [Error] PDFファイルの上書き保存に失敗 ({pdf_path_str}): {e_move}")
            if Path(temp_output_path).is_file():
                try:
                    Path(temp_output_path).unlink()
                except Exception:
                    pass
            raise Exception(f"OCR後のPDFファイル上書きに失敗: {e_move}")


# ==============================================================================
# 画像 -> PDF 変換関数
# ==============================================================================

def convert_image_to_pdf(image_path: Path, output_pdf_path: Path) -> None:
    """
    単一の画像ファイル（png, jpgなど）を1ページのPDFに変換します。
    PyMuPDF (fitz) を使用します。

    Args:
        image_path (Path): 入力画像ファイルのパス。
        output_pdf_path (Path): 出力先PDFファイルのパス。

    Raises:
        Exception: 変換プロセス中にエラーが発生した場合。
    """
    img_doc = None
    pdf_doc = None
    try:
        img_doc = fitz.open(image_path)
        # 画像をPDFのバイトデータに変換
        pdf_bytes = img_doc.convert_to_pdf()
        if not pdf_bytes:
            raise Exception("画像からPDFへの変換に失敗しました。")

        # 新しいPDFドキュメントをバイトデータから作成
        pdf_doc = fitz.open("pdf", pdf_bytes)
        pdf_doc.save(str(output_pdf_path))

    except Exception as e:
        print(f"エラー: {image_path.name} のPDF変換に失敗しました。 {e}", file=sys.stderr)
        raise
    finally:
        if img_doc:
            img_doc.close()
        if pdf_doc:
            pdf_doc.close()


def convert_pil_image_to_pdf(
        pil_image: Image.Image,
        output_pdf_path: Path
) -> None:
    """
    Pillow (PIL) の Image オブジェクトを1ページのPDFに変換します。
    （クリップボードからの画像貼り付け用）

    Args:
        pil_image (Image.Image): Pillow イメージオブジェクト。
        output_pdf_path (Path): 出力先PDFファイルのパス。

    Raises:
        Exception: 変換プロセス中にエラーが発生した場合。
    """
    pdf_doc = None
    img_bytes_io = None
    img_doc = None

    try:
        # 1. PillowイメージをPNG形式でメモリ上のバイトデータに変換
        img_bytes_io = io.BytesIO()
        if pil_image.mode == 'RGBA':
            pil_image = pil_image.convert('RGB')  # 透過情報を除去

        pil_image.save(img_bytes_io, format='PNG')
        img_bytes = img_bytes_io.getvalue()

        # 2. メモリ上のPNGデータをfitzオブジェクトとして開く
        img_doc = fitz.open("png", img_bytes)

        # 3. 画像をPDFのバイトデータに変換
        pdf_bytes = img_doc.convert_to_pdf()
        if not pdf_bytes:
            raise Exception("画像(PIL)からPDFへの変換に失敗しました。")

        # 4. 新しいPDFドキュメントをバイトデータから作成・保存
        pdf_doc = fitz.open("pdf", pdf_bytes)
        pdf_doc.save(str(output_pdf_path))

    except Exception as e:
        print(f"エラー: クリップボード画像のPDF変換に失敗しました。 {e}", file=sys.stderr)
        raise
    finally:
        if img_bytes_io:
            img_bytes_io.close()
        if img_doc:
            img_doc.close()
        if pdf_doc:
            pdf_doc.close()


# ==============================================================================
# Markdown -> PDF 変換関数
# ==============================================================================
def convert_markdown_to_pdf(
    markdown_path: Path,
    output_pdf_path: Path,
    paper_size_str: str = "A4",
) -> None:
    """
    Pandoc (MD->HTML) と Playwright (HTML->PDF) を使用して .md を PDF に変換します。
    Pandoc と Playwright (chromium) がインストールされている必要があります。
    変換前に <details> を <details open> に置換します。

    Args:
        markdown_path (Path): 入力Markdownファイルのパス。
        output_pdf_path (Path): 出力先PDFファイルのパス。
        paper_size_str (str): "A4" または "A5" (config.iniの値)。
        latex_font_name (str): (この関数では未使用)
    """

    # 一時ファイル用のパスを定義
    temp_modified_md_path = output_pdf_path.with_suffix(".temp.md")
    temp_html_path = output_pdf_path.with_suffix(".temp.html")

    # --- ステップ 1: <details> を <details open> に置換 ---
    try:
        # 元のMarkdownファイルを読み込む
        with open(markdown_path, 'r', encoding='utf-8') as f:
            md_content = f.read()

        # <details> タグを <details open> に置換 (大文字小文字を区別しない)
        # 既に 'open' があっても 'open open' にならないよう、単純な置換を避ける
        # '<details' (末尾スペースなし) または '<details ' (末尾スペースあり) を検索
        modified_md_content = re.sub(
            r"<details(?![^>]*\bopen\b)",  # 'open'属性をまだ持たない<details>タグ
            "<details open",              # '<details open' に置換
            md_content,
            flags=re.IGNORECASE          # 大文字小文字を無視
        )

        # 置換後の内容を一時的な .md ファイルに書き出す
        with open(temp_modified_md_path, 'w', encoding='utf-8') as f:
            f.write(modified_md_content)

    except Exception as e:
        raise Exception(f"Markdownの前処理(<details>置換)に失敗しました: {e}")

    # --- ステップ 2: Pandoc で Markdown を HTML (一時ファイル) に変換 ---
    input_format = "gfm"

    pandoc_cmd = [
        "pandoc",
        "--from", input_format,
        str(temp_modified_md_path),  # [変更] 置換後の一時MDファイルを使用
        "-s",                        # スタンドアロン (HTMLヘッダ等を含む)
        "--embed-resources",         # 画像などをHTMLに埋め込む
        "--mathml",                  # 数式をMathML (HTML互換) に変換
        "--to", "html5",
        "-o", str(temp_html_path)
    ]
    print(f"  [Info] Pandoc (MD->HTML) 実行: {' '.join(pandoc_cmd)}")

    try:
        subprocess.run(
            pandoc_cmd,
            capture_output=True,
            text=True,
            encoding='utf-8',
            errors='ignore',
            check=True
        )
    except FileNotFoundError:
        # finallyブロックで一時MDファイルが削除されるよう、エラーを再送出
        raise Exception(
            "Pandoc が見つかりません。\n" +
            "Markdown連携には Pandoc のインストールとPATH設定が必要です。"
        )
    except subprocess.CalledProcessError as e:
        error_details = f"STDOUT:\n{e.stdout}\n\nSTDERR:\n{e.stderr}"
        # finallyブロックで一時MDファイルが削除されるよう、エラーを再送出
        raise Exception(
            f"PandocでのHTML変換に失敗しました (ReturnCode {e.returncode}):\n"
            f"{error_details}"
        )

    # --- ステップ 2: Playwright で HTML を PDF に変換 ---
    playwright_paper_format = paper_size_str.upper()

    pw_instance = None
    browser = None
    page = None
    try:
        print(f"  [Info] Playwright (HTML->PDF) 実行: {temp_html_path.name}")
        pw_instance = sync_playwright().start()
        browser = pw_instance.chromium.launch()
        page = browser.new_page()

        page.goto(temp_html_path.as_uri(), wait_until='networkidle')

        page.pdf(
            path=str(output_pdf_path),
            format=playwright_paper_format,
            print_background=True,
            margin={
                'top': '1cm', 'bottom': '1cm',
                'left': '1cm', 'right': '1cm'
            }
        )
        print(f"  [Info] Playwright PDF変換完了: {output_pdf_path.name}")

    except PlaywrightError as e:
        # finallyブロックで一時MDファイルが削除されるよう、エラーを再送出
        raise Exception(
            "Playwright (Chromium) でのHTML->PDF変換に失敗しました。\n" +
            "Install.bat を実行して Playwright が正しくインストールされているか確認してください。\n" +
            f"エラー: {e}"
        )
    except Exception as e:
        # finallyブロックで一時MDファイルが削除されるよう、エラーを再送出
        raise Exception(f"Playwright PDF変換中の予期せぬエラー: {e}")
    finally:
        # Playwrightセッションを必ず閉じる
        if page:
            page.close()
        if browser:
            browser.close()
        if pw_instance:
            pw_instance.stop()

        # [変更] 一時HTMLファイル と 一時MDファイル の両方を削除
        if temp_html_path.is_file():
            try:
                temp_html_path.unlink()
            except Exception as e_del:
                print(f"  [Warn] 一時HTMLファイルの削除に失敗: {e_del}")

        if temp_modified_md_path.is_file():
            try:
                temp_modified_md_path.unlink()
            except Exception as e_del:
                print(
                    f"  [Warn] 一時MDファイル({temp_modified_md_path.name})の削除に失敗: "
                    f"{e_del}"
                )

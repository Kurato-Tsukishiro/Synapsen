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
import re
import subprocess
import qrcode
import json

from playwright.sync_api import sync_playwright, Error as PlaywrightError

import logging
logger = logging.getLogger(__name__)

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
        logger.error(f"警告: 16進数カラーコード '{hex_color}' の変換に失敗: {e}")
        return None


def add_metadata_to_clip(
    pdf_path_str: str,
    font_path: str,
    paper_width: float,
    paper_height: float,
    key_rect_tuple: tuple,
    index_key_to_embed: str,
    text_color: tuple | None,
    comment_to_embed: str,
    sist_string_formal: str | None = None,
    sist_string_readable: str | None = None,
    base_name: str | None = None,
    cited_keys_list: list[str] | None = None,
    refs_qr_size_pt: int = 75
) -> None:
    """
    Playwrightで生成されたPDFに対し、
    1ページ目に IndexKey (テキスト+QR) を、最終ページ（新規追加）に コメントと書誌情報及び引用Key(QR) を書き込みます。
    """

    # --- 埋め込む情報が何もなければ、処理をスキップ ---
    if (
        not index_key_to_embed and
        not comment_to_embed and
        not sist_string_formal and
        not cited_keys_list and
        not base_name
    ):
        logger.info(
            f"埋め込むメタデータがないためスキップ: {Path(pdf_path_str).name}",
            extra={'sensitive': True}
        )
        return

    doc = None
    try:
        doc = fitz.open(pdf_path_str)
        if len(doc) == 0:
            logger.error(
                f"メタデータ埋め込みスキップ: ページが存在しません {pdf_path_str}",
                extra={'sensitive': True}
            )
            return

        key_rect = fitz.Rect(key_rect_tuple)
        font_alias = "embed_font"

        # --- 1. 1ページ目に IndexKey (QR + テキスト) を描画 ---
        if index_key_to_embed or base_name:
            page1 = doc[0]
            try:
                page1.insert_font(fontname=font_alias, fontfile=font_path)
            except Exception as e:
                logger.warning(f"フォント埋め込み警告 (Page 1): {e}")

        # ============================================================
        # A. QRコードの生成と描画 (Page 1: cpk と key のみ)
        # ============================================================
            try:
                # --- QRコードに埋め込むJSONデータを構築 (cpk と key のみ) ---
                qr_data = {"cpk": index_key_to_embed}
                auto_generated_key = ""

                if base_name:
                    # base_name をパースして Ersteller と同じ "key" のみ生成
                    match = re.match(
                        r"(\d{8})_(?:(\d{4,6})_)?(.+)",
                        base_name,
                        re.IGNORECASE)

                    if match:
                        date_str, time_val, _ = match.groups()

                        if time_val:
                            time_str = time_val.ljust(6, '0')
                        else:
                            time_str = "999999"

                        if time_str != "999999":
                            key_time = time_str
                        else:
                            key_time = "000000"
                        auto_generated_key = date_str + key_time

                        qr_data["key"] = auto_generated_key

                # JSON文字列に変換
                qr_data_str = json.dumps(qr_data, ensure_ascii=False)
                # --- 構築完了 ---

                qr = qrcode.QRCode(box_size=2, border=0)
                qr.add_data(qr_data_str)
                qr.make(fit=True)
                qr_img = qr.make_image(fill_color="black", back_color="white")

                # バイト列に変換
                img_byte_arr = io.BytesIO()
                qr_img.save(img_byte_arr, format='PNG')
                img_bytes = img_byte_arr.getvalue()

                # (Page 1 QR のレイアウト計算 ... 変更なし)
                qr_size = 35
                qr_margin = 5
                qr_x = key_rect.x0 + qr_margin
                qr_y = key_rect.y0
                qr_rect = fitz.Rect(qr_x, qr_y, qr_x + qr_size, qr_y + qr_size)

                page1.insert_image(qr_rect, stream=img_bytes)

            except Exception as e:
                logger.error(f"QRコード生成エラー (Page 1): {e}")

            # ============================================================
            # B. テキストの描画 (Index Key がある場合のみ)
            # ============================================================
            # Index Key がある場合のみテキストを描画
            shape1 = page1.new_shape()
            text_rect = fitz.Rect(
                qr_rect.x1 + qr_margin, key_rect.y0,
                key_rect.x1, key_rect.y1
            )
            if index_key_to_embed:
                shape1.insert_textbox(
                    text_rect, index_key_to_embed, fontname=font_alias,
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
                logger.warning(f"フォント埋め込み警告 (Last Page): {e}")

            shape_last = page_last.new_shape()

            # --- 描画座標の計算 (embed_metadata_as_cover_page と同様) ---
            info_rect_y_start = key_rect.y1 + 10
            info_rect_y_end = page_last.rect.height - 30
            current_y_pos = info_rect_y_start
            x0 = key_rect.x0

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
                current_y_pos = actual_sist_y1 + 10

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
                current_y_pos = actual_readable_y1 + 40
            else:
                current_y_pos = info_rect_y_start + 40

            comment_rect = fitz.Rect(
                x0, current_y_pos,
                page_last.rect.width - 50, info_rect_y_end
            )
            if comment_to_embed:
                shape_last.insert_textbox(
                    comment_rect, f"コメント:\n{comment_to_embed}",
                    fontname=font_alias, fontsize=9, align=0
                )

            shape_last.commit()

        # 最終ページに「引用Key専用QRコード」を描画
        if cited_keys_list:
            # (page_last がまだ定義されていない場合 = コメント等が空だった場合)
            if 'page_last' not in locals():
                page_last = doc.new_page(
                    pno=len(doc), width=paper_width, height=paper_height
                )

            try:
                # 1. Prepare JSON (refs only)
                qr_data_refs = {"refs": cited_keys_list}
                qr_data_str_refs = json.dumps(qr_data_refs, ensure_ascii=False)

                # 2. Generate QR image (Larger size)
                qr_refs = qrcode.QRCode(
                    box_size=4,
                    border=1
                )
                qr_refs.add_data(qr_data_str_refs)
                qr_refs.make(fit=True)
                qr_img_refs = qr_refs.make_image(
                    fill_color="black", back_color="white")

                # 3. Get image bytes
                img_byte_arr_refs = io.BytesIO()
                qr_img_refs.save(img_byte_arr_refs, format='PNG')
                img_bytes_refs = img_byte_arr_refs.getvalue()

                # 4. Define position (Bottom-Right)
                qr_size_refs = refs_qr_size_pt  # Default: 75x75 pt (約 2.6 cm)
                margin_refs = 30                # 右下からのマージン
                qr_x_refs = page_last.rect.width - qr_size_refs - margin_refs
                qr_y_refs = page_last.rect.height - qr_size_refs - margin_refs
                qr_rect_refs = fitz.Rect(
                    qr_x_refs, qr_y_refs,
                    qr_x_refs + qr_size_refs, qr_y_refs + qr_size_refs
                )

                # 5. Insert QR image
                page_last.insert_image(qr_rect_refs, stream=img_bytes_refs)
                logger.info(
                    f"引用Key専用QRコードを最終ページに埋め込みました: {Path(pdf_path_str).name}",
                    extra={'sensitive': True}
                )

            except Exception as e:
                logger.error(f"最終ページの引用Key用QRコード生成エラー: {e}")

        # 変更を上書き保存
        doc.saveIncr()

    except Exception as e:
        logger.error(
            f"Webクリップへのメタデータ埋め込み中にエラー ({pdf_path_str}): {e}",
            extra={'sensitive': True}
        )
        raise
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
            logger.warning(
                f"暗号化されたPDFはスキップします: {Path(input_path).name}",
                extra={'sensitive': True}
            )
            return  # 暗号化ファイルは処理せず終了

        font_name_in_pdf = "synapsen-embed-font"  # PDF内部で使うフォントのエイリアス名

        for page in doc:
            try:
                page.insert_font(fontname=font_name_in_pdf, fontfile=font_path)
            except Exception as e:
                logger.info(f"Font insertion issue ({e}). Continuing.")

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
        logger.error(
            f"フラット化処理中にエラー ({input_path}): {e}",
            extra={'sensitive': True}
        )
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
                logger.warning(
                    f"Skipping empty or invalid page in {input_path}",
                    extra={'sensitive': True}
                )
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
        logger.error(
            f"正規化処理中にエラー ({input_path}): {e}",
            extra={'sensitive': True}
        )
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
            logger.info(
                f"暗号化されたPDFはスキップします: {Path(pdf_path_str).name}",
                extra={'sensitive': True}
            )
            return

        # 1. 高速なテキスト抽出を試みる
        meaningful_text_threshold = 10

        if not enable_tesseract:
            logger.info("Tesseract OCR は無効です。スキップします。")
            return

        logger.info(
            "Tesseract OCR を実行し、テキストを埋め込みます: "
            f"{Path(pdf_path_str).name}"
        )

        pages_processed_count = 0

        for page_num, page in enumerate(doc):

            # 1. ページごとに既存テキストをチェック
            page_text = page.get_text("text", sort=True).strip()
            if len(page_text) > meaningful_text_threshold:
                logger.info(
                    f"Page {page_num + 1} には既存テキストがあるためスキップ。",
                    extra={'sensitive': True}
                )
                continue

            # --- 既存テキストがないページのみ、以下を実行 ---
            logger.info(
                f"Tesseract OCR を実行中 (Page {page_num + 1})...",
                extra={'sensitive': True}
            )
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
                    logger.warning(
                        "TesseractがTSVデータを返しませんでした "
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
                    logger.info(
                        "Tesseract OCR は実行されましたが、"
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
                logger.warning(
                    f"Tesseract OCRエラー (Page {page_num + 1}): {ocr_err}")
                continue

        if pages_processed_count == 0:
            logger.info("OCRが実行されたページはありませんでした。")
            return

        # 11. 変更を「一時ファイル」に保存
        doc.save(
            temp_output_path,
            garbage=4,
            deflate=True,
            encryption=fitz.PDF_ENCRYPT_NONE
        )
        logger.info(
            f"テキスト埋め込み完了 (一時ファイル): {Path(temp_output_path).name}",
            extra={'sensitive': True}
        )

    except Exception as e:
        logger.error(
            f"PDFテキスト埋め込み処理中にエラー ({pdf_path_str}): {e}",
            extra={'sensitive': True}
        )
        if Path(temp_output_path).is_file():
            try:
                Path(temp_output_path).unlink()  # エラー時は一時ファイルを削除
            except Exception as e_del:
                logger.warning(f"エラー発生後の一時ファイル削除に失敗: {e_del}")
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
            logger.info(
                f"元ファイルに上書き完了: {Path(pdf_path_str).name}",
                extra={'sensitive': True}
            )
        except Exception as e_move:
            logger.error(
                f"PDFファイルの上書き保存に失敗 ({pdf_path_str}): {e_move}",
                extra={'sensitive': True}
            )
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
        logger.error(
            f"{image_path.name} のPDF変換に失敗しました。 {e}",
            extra={'sensitive': True}
        )
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
        logger.error(f"クリップボード画像のPDF変換に失敗しました。 {e}")
        raise
    finally:
        if img_bytes_io:
            img_bytes_io.close()
        if img_doc:
            img_doc.close()
        if pdf_doc:
            pdf_doc.close()


# ==============================================================================
# Markdown(他) -> PDF 変換関数
# ==============================================================================
# マッピング辞書を定義
PANDOC_INPUT_FORMATS = {
    ".md": "gfm",
    ".txt": "plain",
    ".rtf": "rtf",
    ".docx": "docx",
    ".odt": "odt"
}


def convert_document_to_pdf(
    input_path: Path,
    output_pdf_path: Path,
    paper_size_str: str = "A4",
) -> None:
    """
    Pandoc (MD, TXT, DOCX等) と Playwright (HTML->PDF) を使用して ドキュメント を PDF に変換します。
    Pandoc と Playwright (chromium) がインストールされている必要があります。
    変換前に <details> を <details open> に置換します。

    Args:
        input_path (Path): 入力ファイルのパス。
        output_pdf_path (Path): 出力先PDFファイルのパス。
        paper_size_str (str): "A4" または "A5" (config.iniの値)。
        latex_font_name (str): (この関数では未使用)
    """

    # 一時ファイル用のパスを定義
    temp_modified_content_path = output_pdf_path.with_suffix(".temp.modified")
    temp_html_path = output_pdf_path.with_suffix(".temp.html")

    file_suffix = input_path.suffix.lower()

    # --- ステップ 1: 前処理 (フォーマットごとに行う) ---
    try:
        with open(input_path, 'r', encoding='utf-8') as f:
            content = f.read()

        # Markdownの場合のみ <details> を置換
        if file_suffix == ".md":
            modified_content = re.sub(
                r"<details(?![^>]*\bopen\b)",
                "<details open",
                content,
                flags=re.IGNORECASE
            )
        # テキストファイルの場合、改行を維持するために <pre> タグで囲む
        elif file_suffix == ".txt":
            import html
            escaped_content = html.escape(content)
            # preタグで囲み、CSSでフォントと言語を指定 (font-familyはシステムのsans-serifに依存させます)
            modified_content = (
                "<pre style='white-space: pre-wrap; "
                "font-family: sans-serif;'>"
                f"{escaped_content}</pre>"
            )
        else:
            modified_content = content  # DOCXなどはそのままPandocに渡す

        # 前処理が不要な形式（docxなど）と、処理済みの内容を一時ファイルに書き出す
        if file_suffix in [".docx", ".rtf", ".odt"]:
            # バイナリファイルをコピー
            shutil.copy2(input_path, temp_modified_content_path)
        else:
            # テキストベースのファイルを書き出し
            with open(temp_modified_content_path, 'w', encoding='utf-8') as f:
                f.write(modified_content)

    except Exception as e:
        # DOCXなどは 'utf-8' で読めないため、バイナリとして扱う
        if file_suffix in PANDOC_INPUT_FORMATS and file_suffix not in [
                ".md", ".txt"]:
            try:
                shutil.copy2(input_path, temp_modified_content_path)
            except Exception as copy_e:
                raise Exception(f"ドキュメントの前処理（コピー）に失敗しました: {copy_e}")
        else:
            raise Exception(f"ドキュメントの前処理（読み込み）に失敗しました: {e}")

    # --- ステップ 2: Pandoc で HTML (一時ファイル) に変換 ---
    # マッピング辞書から入力フォーマットを取得
    input_format = PANDOC_INPUT_FORMATS.get(file_suffix, "gfm")  # 不明な場合はgfm扱い

    pandoc_cmd = [
        "pandoc",
        "--from", input_format,
        str(temp_modified_content_path),  # 処理後の一時ファイルを使用
        "-s",
        "--embed-resources",
        "--mathml",
        "--to", "html5",
        "-o", str(temp_html_path)
    ]
    logger.info(
        f"Pandoc (MD->HTML) 実行: {' '.join(pandoc_cmd)}",
        extra={'sensitive': True}
    )

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
        logger.info(
            f"Playwright (HTML->PDF) 実行: {temp_html_path.name}",
            extra={'sensitive': True}
        )
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
        logger.info(
            f"Playwright PDF変換完了: {output_pdf_path.name}",
            extra={'sensitive': True}
        )

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

        # 一時HTMLファイル と 一時処理ファイルを両方削除
        if temp_html_path.is_file():
            try:
                temp_html_path.unlink()
            except Exception as e_del:
                logger.warning(f"一時HTMLファイルの削除に失敗: {e_del}")

        if temp_modified_content_path.is_file():
            try:
                temp_modified_content_path.unlink()
            except Exception as e_del:
                logger.warning(
                    f"一時処理ファイル({temp_modified_content_path.name})の削除に失敗: "
                    f"{e_del}",
                    extra={'sensitive': True}
                )

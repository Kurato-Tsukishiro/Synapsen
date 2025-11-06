import fitz  # PyMuPDF
from pypdf import PdfReader, PdfWriter, Transformation
from pathlib import Path
import pytesseract
from PIL import Image
import io
import pandas as pd
from pytesseract import Output
import csv
import shutil  # <--- ファイル操作（上書きリネーム）のために必要
import sys

# ==============================================================================
# 定数定義
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


def high_fidelity_flatten(input_path: str, output_path: str, font_path: str):
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
    """
    if not Path(font_path).is_file():
        raise FileNotFoundError(f"指定されたフォントファイルが見つかりません: {font_path}")

    doc = fitz.open(input_path)
    font_name_in_pdf = "notosans-jp"  # PDF内部で使うフォントのエイリアス名

    for page in doc:
        # ページにカスタムフォントを登録・埋め込む
        try:
            page.insert_font(fontname=font_name_in_pdf, fontfile=font_path)
        except Exception as e:
            # 既に登録されている場合などがあるので、エラーが出ても処理を続行
            print(f"Info: Font insertion issue ({e}). Continuing.")

        # フォームウィジェットを処理
        for widget in page.widgets():
            # テキストフィールドまたはコンボボックスで、値が存在する場合
            if widget.field_type in (
                fitz.PDF_WIDGET_TYPE_TEXT,
                fitz.PDF_WIDGET_TYPE_COMBOBOX
            ) and widget.field_value:
                # フィールドの値（テキスト）をページに直接描画
                page.insert_textbox(
                    widget.rect,  # フィールドと同じ位置・サイズ
                    widget.field_value,
                    fontname=font_name_in_pdf,
                    fontsize=widget.text_fontsize or 10,
                    color=widget.text_color or (0, 0, 0),
                )

            # 元のインタラクティブなウィジェットを削除
            page.delete_widget(widget)

    # PDFを保存 (ガベージコレクション、圧縮を有効化)
    doc.save(output_path, garbage=4, deflate=True)
    doc.close()


def normalize_pdf_to_papersize(
        input_path: str,
        output_path: str,
        paper_width: float,
        paper_height: float
        ):
    """
    PDFの全ページを、指定された用紙サイズの中央にリサイズ・配置します。

    マージン領域を考慮し、コンテンツがその領域内に収まるように
    アスペクト比を維持してスケーリングおよび中央配置を行います。

    Args:
        input_path (str): 入力PDFファイル（通常はフラット化済み）のパス。
        output_path (str): 正規化された出力PDFファイルのパス。
        paper_width (float): ターゲットの用紙幅 (ポイント単位)。
        paper_height (float): ターゲットの用紙高 (ポイント単位)。
    """
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

        # 描画可能領域 (drawable_width, drawable_height) を使用
        scale = min(
            drawable_width / original_width,
            drawable_height / original_height
        )

        # 描画可能領域内で中央に配置
        tx = LEFT_MARGIN + (drawable_width - original_width * scale) / 2
        ty = BOTTOM_MARGIN + (drawable_height - original_height * scale) / 2

        transform =\
            Transformation().scale(sx=scale, sy=scale).translate(tx=tx, ty=ty)

        template_page.merge_transformed_page(content_page, transform)

    with open(output_path, "wb") as f:
        writer.write(f)


def embed_ocr_text_in_pdf(
    pdf_path_str: str,
    enable_tesseract: bool,
    font_path: str,
    lang='jpn+jpn_vert'
):
    """
    PDFを解析し、必要に応じてTesseract OCRを実行し、
    結果を「透明なテキストレイヤー」としてPDF自体に埋め込む。

    Args:
        pdf_path_str (str): 処理対象のPDFファイルパス（読み書きされる）。
        enable_tesseract (bool): Tesseract OCR (低速) を実行するかどうか。
        font_path (str): 埋め込む日本語フォントファイルのパス。
        lang (str): Tesseractが使用する言語。
    """
    doc = None
    # 一時ファイルへの保存パスを定義
    temp_output_path = pdf_path_str + "._temp_ocr.pdf"

    # 埋め込みに使うフォント名を定義
    OCR_FONT_NAME = "synapsen_ocr_font"

    try:
        doc = fitz.open(pdf_path_str)
        if doc.is_encrypted:
            print(f"  [Warn]暗号化されたPDFはスキップします: {Path(pdf_path_str).name}")
            return

        meaningful_text_threshold = 10
        has_meaningful_text = False

        # 2. 高速なテキスト抽出を試みる
        for page in doc:
            page_text = page.get_text("text", sort=True).strip()
            if len(page_text) > meaningful_text_threshold:
                has_meaningful_text = True
                break

        # 3. テキストが取得できたか確認
        if has_meaningful_text:
            print(
                "  [Info] 既存のテキストレイヤーが存在するためスキップ: " +
                f"{Path(pdf_path_str).name}"
                )
            return

        # 4. Tesseract OCR が無効化されているかチェック
        if not enable_tesseract:
            print(
                "  [Info] 画像のみのPDFで、Tesseract OCR は無効です。" +
                f"スキップ: {Path(pdf_path_str).name}"
                )
            return

        # 5. 既存テキストが取れず、Tesseract が有効な場合のみ実行
        print(
            "  [Info] Tesseract OCR を実行し、テキストを埋め込みます: " +
            f"{Path(pdf_path_str).name}"
            )

        for page_num, page in enumerate(doc):
            try:
                # 6. 高解像度 (DPI=300) でページを画像(Pixmap)にレンダリング
                pix = page.get_pixmap(dpi=300)
                img_data = pix.tobytes("png")
                img = Image.open(io.BytesIO(img_data))

                # 7. Tesseract OCR の実行 (TSVデータとして取得)
                #    lang='jpn+jpn_vert' が正しく渡される
                tsv_data = pytesseract.image_to_data(
                    img, lang=lang, output_type=Output.STRING
                )

                if tsv_data is None or len(tsv_data.strip()) == 0:
                    print(
                        "  [Warn] TesseractがTSVデータを返しませんでした " +
                        f"(Page {page_num + 1})"
                        )
                    continue

                # 8. TesseractのTSVデータを解析
                try:
                    df = pd.read_csv(
                        io.StringIO(tsv_data),
                        sep='\t', quoting=csv.QUOTE_NONE, on_bad_lines='skip')
                    df = df.dropna(subset=['conf', 'text'])
                    df = df[df['conf'] > 30]  # 信頼度が低いものは除外

                    if df.empty:
                        continue

                    # 9. ページに日本語フォント (config.ini の font_path) を登録
                    try:
                        # font_path (例: Noto Sans) を 'OCR_FONT_NAME' として登録
                        page.insert_font(
                            fontname=OCR_FONT_NAME, fontfile=font_path)
                    except Exception:
                        pass  # 既に登録済みなどのエラーは無視

                    # 10. 登録したフォント名を使って、透明テキストを挿入
                    for _, row in df.iterrows():
                        x0 = row['left']
                        y0 = row['top']
                        w = row['width']
                        h = row['height']
                        scale = 72 / 300  # DPI=300 -> 72 DPI (ポイント) に座標を戻す
                        rect = fitz.Rect(
                            x0 * scale, y0 * scale,
                            (x0 + w) * scale, (y0 + h) * scale
                            )
                        fs = max(h * scale * 0.8, 6.0)  # フォントサイズ

                        page.insert_text(
                            rect.bottom_left,
                            str(row['text']),
                            fontname=OCR_FONT_NAME,
                            fontsize=fs,
                            render_mode=3,  # 3 = 透明
                            rotate=0
                        )

                except Exception as e_embed:
                    print("  [Warn] テキストの埋め込みに失敗。")
                    print(f"  [Debug] エラー詳細: {e_embed}")

            except pytesseract.TesseractNotFoundError:
                raise Exception("Tesseract-OCRが見つかりません。")
            except Exception as ocr_err:
                print(
                    "  [Warn] Tesseract OCRエラー " +
                    f"(Page {page_num + 1}): {ocr_err}"
                    )
                continue

        # 11. 変更を「一時ファイル」に保存 (エラーの出ない引数のみ)
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
                Path(temp_output_path).unlink()
            except Exception as e_del:
                print(f"  [Warn] エラー発生後の一時ファイル削除に失敗: {e_del}")
    finally:
        if doc:
            doc.close()

        # 12. ファイルのリネーム（上書き）
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


# ==============================================================================
# 画像 -> PDF 変換関数
# ==============================================================================
def convert_image_to_pdf(image_path: Path, output_pdf_path: Path):
    """
    単一の画像ファイル（png, jpgなど）を1ページのPDFに変換します。
    PyMuPDFは多くの画像形式の読み込みに標準で対応しています。

    Args:
        image_path (Path): 入力画像ファイルのパス。
        output_pdf_path (Path): 出力先PDFファイルのパス。

    Returns:
        bool: 変換が成功したかどうか。

    Raises:
        Exception: 変換プロセス中にエラーが発生した場合。
    """
    img_doc = None
    pdf_doc = None
    try:
        # 1. 画像ファイルをfitzオブジェクトとして開く
        img_doc = fitz.open(image_path)

        # 2. 画像をPDFのバイトデータに変換する
        #    [修正] rect 引数を削除 (エラーログ対応)
        pdf_bytes = img_doc.convert_to_pdf()

        if not pdf_bytes:
            raise Exception("画像からPDFへの変換に失敗しました。")

        # 3. 新しいPDFドキュメントをバイトデータから作成
        pdf_doc = fitz.open("pdf", pdf_bytes)

        # 4. PDFとして保存
        pdf_doc.save(str(output_pdf_path))

        return True

    except Exception as e:
        # エラーを捕捉し、呼び出し元に伝える
        print(f"エラー: {image_path.name} のPDF変換に失敗しました。 {e}", file=sys.stderr)
        raise  # エラーを再送出し、main.py側で処理できるようにする

    finally:
        # ドキュメントを確実に閉じる
        if img_doc:
            img_doc.close()
        if pdf_doc:
            pdf_doc.close()

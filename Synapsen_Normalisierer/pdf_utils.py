import fitz  # PyMuPDF
from pypdf import PdfReader, PdfWriter, Transformation
from pathlib import Path
import pytesseract
from PIL import Image
import io

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


def perform_ocr_on_pdf(
    pdf_path_str: str,
    output_txt_path_str: str,
    enable_tesseract: bool,
    lang='jpn+jpn_vert'
):
    """
    PDFからテキストを抽出する。

    1. まず高速な組み込みテキスト抽出を試みる (既にOCR/テキストがある場合)。
    2. 1.でテキストが取れなかった場合、enable_tesseract が True のみ、
       Tesseract OCR (低速) を実行する。

    Args:
        pdf_path_str (str): OCR対象のPDFファイルパス（フラット化済みのもの）。
        output_txt_path_str (str): 抽出テキストの保存先パス。
        enable_tesseract (bool): Tesseract OCR (低速) を実行するかどうか。
        lang (str): Tesseractが使用する言語。
    """
    full_text = ""
    doc = None
    try:
        # 1. PyMuPDF (fitz) でPDFを開く
        doc = fitz.open(pdf_path_str)

        # 2. 高速なテキスト抽出を試みる
        for page_num in range(len(doc)):
            page = doc.load_page(page_num)
            # "text" モードでテキストを抽出
            page_text = page.get_text("text", sort=True)
            if page_text:
                full_text += page_text + "\n\n--- Page Break ---\n\n"

        # 3. テキストが取得できたか確認
        #    (空白文字を除いて10文字以上ある場合を「意味のあるテキスト」とみなす)
        meaningful_text_threshold = 10
        if len(full_text.strip()) > meaningful_text_threshold:
            # 取得できた場合 (既にテキストレイヤーが存在した)
            print(f"  [Info] 既存のテキストを抽出しました: {Path(pdf_path_str).name}")
            with open(output_txt_path_str, "w", encoding="utf-8") as f:
                f.write(full_text)
            return  # ★ finallyが呼ばれてからreturnされる

        # 4. Tesseract OCR が無効化されているかチェック
        if not enable_tesseract:
            print(f"  [Info] 既存テキストが見つからず、Tesseract OCR は無効です。スキップします: {
                  Path(pdf_path_str).name}")
            # Ersteller がカラムを見失わないよう、空のテキストファイルを作成する
            with open(output_txt_path_str, "w", encoding="utf-8") as f:
                f.write("")  # 空のファイルを作成
            return  # ★ finallyが呼ばれてからreturnされる

        # 5. 既存テキストが取れず、Tesseract が有効な場合のみ実行
        print(f"  [Info] 既存テキストが見つかりません。Tesseract OCR を実行します: {
              Path(pdf_path_str).name}")
        full_text = ""  # テキストをリセット
        for page_num in range(len(doc)):
            page = doc.load_page(page_num)

            # 6. 高解像度 (DPI=300) でページを画像(Pixmap)にレンダリング
            pix = page.get_pixmap(dpi=300)
            img_data = pix.tobytes("png")

            # 7. PIL (Pillow) を使って画像データをTesseractが読める形式に変換
            img = Image.open(io.BytesIO(img_data))

            # 8. Tesseract OCR の実行
            try:
                page_text = pytesseract.image_to_string(img, lang=lang)
                full_text += page_text + "\n\n--- Page Break ---\n\n"
            except pytesseract.TesseractNotFoundError:
                raise Exception(
                    "Tesseract-OCRが見つかりません。" +
                    "Tesseract-OCRをインストールし、環境変数PATHを設定してください。"
                )
            except Exception as ocr_err:
                print(
                    "  [Warn] Tesseract OCRエラー " +
                    f"(Page {page_num + 1}): {ocr_err}"
                )
                continue

        # 9. 抽出したテキストをファイルに保存
        with open(output_txt_path_str, "w", encoding="utf-8") as f:
            f.write(full_text)

    except Exception as e:
        print(f"  [Error] OCR処理中にエラーが発生 ({pdf_path_str}): {e}")
        with open(output_txt_path_str, "w", encoding="utf-8") as f:
            f.write(f"OCR処理中にエラーが発生しました: {e}")
    finally:
        # ★ どのような場合でも、ここでdocが安全に閉じられる
        if doc:
            doc.close()

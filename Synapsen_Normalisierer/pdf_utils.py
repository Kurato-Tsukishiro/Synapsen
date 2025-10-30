import fitz  # PyMuPDF
from pypdf import PdfReader, PdfWriter, Transformation
from pathlib import Path

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


def normalize_pdf_to_papersize(input_path: str, output_path: str, paper_width: float, paper_height: float):
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

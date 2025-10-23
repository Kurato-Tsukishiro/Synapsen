import fitz  # PyMuPDF
from pypdf import PdfReader, PdfWriter, Transformation
from pathlib import Path

# ==============================================================================
# 定数定義
# ==============================================================================

# A4サイズの寸法 (ポイント単位)
A4_WIDTH: float = 595.276
A4_HEIGHT: float = 841.89

# --- LaTeXのgeometry設定に基づくマージン計算（ポイント単位） ---
CM_TO_PT: float = 72 / 2.54

# LaTeX設定値 (cm)
MARGIN_CM: float = 0
HEAD_SEP_CM: float = 1.0

# ポイント単位に変換
MARGIN: float = MARGIN_CM * CM_TO_PT
HEAD_SEP: float = HEAD_SEP_CM * CM_TO_PT

# ヘッダーとフッターを避け、左右の余白をなくす設定
TOP_MARGIN: float = MARGIN + HEAD_SEP  # 上部余白は維持
BOTTOM_MARGIN: float = MARGIN           # 下部余白は維持
LEFT_MARGIN: float = 0                  # 左余白を0に
RIGHT_MARGIN: float = 0                 # 右余白を0に

# コンテンツを描画する領域のサイズ
DRAWABLE_WIDTH: float = A4_WIDTH - LEFT_MARGIN - RIGHT_MARGIN
DRAWABLE_HEIGHT: float = A4_HEIGHT - TOP_MARGIN - BOTTOM_MARGIN
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


def normalize_pdf_to_a4(input_path: str, output_path: str):
    """
    PDFの全ページを、A4縦サイズの中央にリサイズ・配置します。

    グローバル定数 (DRAWABLE_WIDTH, DRAWABLE_HEIGHT, etc.) に基づく
    マージン領域を考慮し、コンテンツがその領域内に収まるように
    アスペクト比を維持してスケーリングおよび中央配置を行います。

    Args:
        input_path (str): 入力PDFファイル（通常はフラット化済み）のパス。
        output_path (str): A4に正規化された出力PDFファイルのパス。
    """
    reader = PdfReader(input_path)
    writer = PdfWriter()

    for content_page in reader.pages:
        # A4縦の白紙ページをテンプレートとして作成
        template_page = writer.add_blank_page(width=A4_WIDTH, height=A4_HEIGHT)

        original_width = float(content_page.mediabox.width)
        original_height = float(content_page.mediabox.height)

        if original_width == 0 or original_height == 0:
            print(f"Skipping empty or invalid page in {input_path}")
            continue

        # 描画可能領域 (DRAWABLE_WIDTH, DRAWABLE_HEIGHT) に
        # 収まるようにスケーリング係数を計算
        scale = min(
            DRAWABLE_WIDTH / original_width,
            DRAWABLE_HEIGHT / original_height
        )

        # 描画可能領域内で中央に配置するための移動量(オフセット)を計算
        # (tx, ty)はページの左下隅からのオフセット
        tx = LEFT_MARGIN + (DRAWABLE_WIDTH - original_width * scale) / 2
        ty = BOTTOM_MARGIN + (DRAWABLE_HEIGHT - original_height * scale) / 2

        # 変換（スケーリングと移動）を定義
        transform =\
            Transformation().scale(sx=scale, sy=scale).translate(tx=tx, ty=ty)

        # 変換を適用して、元のページコンテンツをテンプレートに合成
        template_page.merge_transformed_page(content_page, transform)

    # 処理が完了したPDFを書き出し
    with open(output_path, "wb") as f:
        writer.write(f)

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

# ollamaのインポート (未導入に対応)
try:
    import ollama
except ImportError:
    ollama = None

# Playwrightのインポート (エラーハンドリング付き)
try:
    from playwright.sync_api import sync_playwright, Error as PlaywrightError
except ImportError:
    sync_playwright = None
    PlaywrightError = Exception

import logging

logger = logging.getLogger(__name__)

# ==============================================================================
# 定数定義
# ==============================================================================

CM_TO_PT: float = 72 / 2.54

# Erstellerで行う統合のレイアウトに合わせた安全な余白設定
# (単位: cm)
LAYOUT_MARGINS = {
    "A4": {
        "top": 2.2,
        "bottom": 2.8,
        "left": 1.0,
        "right": 1.0,
    },
    "A5": {"top": 2.0, "bottom": 2.5, "left": 0.8, "right": 0.8},
}

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
        hex_color = hex_color.lstrip("#")
        # 16進数を 0-255 の整数に変換
        r, g, b = tuple(int(hex_color[i : i + 2], 16) for i in (0, 2, 4))
        # 0-1 の浮動小数点数に変換
        return (r / 255.0, g / 255.0, b / 255.0)
    except Exception as e:
        logger.error(f"警告: 16進数カラーコード '{hex_color}' の変換に失敗: {e}")
        return None


def embed_processing_flag(pdf_path_str: str) -> None:
    """
    PDFのメタデータ(Keywords)に 'Synapsen:SkipNormalization' を追記します。
    これにより、次回以降の正規化処理でサイズ変更がスキップされます。
    """
    doc = None
    try:
        doc = fitz.open(pdf_path_str)
        current_metadata = doc.metadata
        keywords = current_metadata.get("keywords", "")

        skip_flag = "Synapsen:SkipNormalization"

        # まだフラグがない場合のみ追記
        if skip_flag not in keywords:
            new_keywords = f"{keywords}; {skip_flag}" if keywords else skip_flag

            # fitzのset_metadataは辞書全体を渡す必要があるためコピーして更新
            new_metadata = current_metadata.copy()
            new_metadata["keywords"] = new_keywords

            doc.set_metadata(new_metadata)

            # 増分保存 (高速かつ安全)
            doc.saveIncr()

    except Exception as e:
        # メタデータ付与に失敗しても処理自体は止めない（ログのみ）
        logger.warning(
            f"Warning: Failed to embed processing flag to {pdf_path_str}: {e}",
            extra={"sensitive": True},
        )
    finally:
        if doc:
            doc.close()


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
    refs_qr_size_pt: int = 75,
    extra_keywords: list[str] | None = None,
) -> None:
    """
    PDFにメタデータ(QR/テキスト)を描画し、同時にPDFプロパティにも情報を埋め込みます。
    処理済みであることを示すキーワードも埋め込みます。
    """

    # --- 埋め込む情報が何もなければ、処理をスキップ ---
    if (
        not index_key_to_embed
        and not comment_to_embed
        and not sist_string_formal
        and not cited_keys_list
        and not base_name
    ):
        logger.info(
            f"埋め込むメタデータがないためスキップ: {Path(pdf_path_str).name}",
            extra={"sensitive": True},
        )
        return

    doc = None
    try:
        doc = fitz.open(pdf_path_str)
        if len(doc) == 0:
            logger.error(
                f"メタデータ埋め込みスキップ: ページが存在しません {pdf_path_str}",
                extra={"sensitive": True},
            )
            return

        # === 1. メタデータ(PDFプロパティ)への埋め込み処理 ===
        # 既存のメタデータを取得
        current_metadata = doc.metadata

        # 既存のキーワードを取得
        keywords = current_metadata.get("keywords", "")

        # キーワード追加ロジック
        keywords_list = [k.strip() for k in keywords.split(";") if k.strip()]

        # 基本フラグ
        skip_flag = "Synapsen:SkipNormalization"
        if skip_flag not in keywords_list:
            keywords_list.append(skip_flag)

        # 追加キーワード (extra_keywords) の結合
        if extra_keywords:
            for kw in extra_keywords:
                if kw not in keywords_list:
                    keywords_list.append(kw)

        new_keywords = "; ".join(keywords_list)

        new_metadata = current_metadata.copy()
        new_metadata["keywords"] = new_keywords

        # ★ カスタムメタデータとして JSON 形式で情報を埋め込む
        # Subject(件名) や Keywords(キーワード) を汚染しすぎないよう、
        # PyMuPDFの機能を使ってカスタムキーを設定します。
        # (注: 一般的なPDFビューアでは見えませんが、Synapsenからは読み取れます)

        meta_info = {}
        if index_key_to_embed:
            meta_info["cpk"] = index_key_to_embed
        if cited_keys_list:
            meta_info["refs"] = cited_keys_list
        if comment_to_embed:
            meta_info["comment"] = comment_to_embed

        # キー生成ロジック
        auto_generated_key = ""
        if base_name:
            match = re.match(r"(\d{8})_(?:(\d{4,6})_)?(.+)", base_name, re.IGNORECASE)
            if match:
                date_str, time_val, _ = match.groups()
                time_str = time_val.ljust(6, "0") if time_val else "999999"
                key_time = time_str if time_str != "999999" else "000000"
                auto_generated_key = date_str + key_time
                meta_info["key"] = auto_generated_key

        if meta_info:
            # Subjectフィールドの末尾に <synapsen>...</synapsen> で囲んで追記
            json_str = json.dumps(meta_info, ensure_ascii=False)
            current_subject = new_metadata.get("subject", "") or ""
            new_metadata["subject"] = (
                f"{current_subject}\n<synapsen>{json_str}</synapsen>"
            )

        # メタデータを適用
        doc.set_metadata(new_metadata)

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
                if base_name and auto_generated_key:
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
                qr_img.save(img_byte_arr, format="PNG")
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
                qr_rect.x1 + qr_margin, key_rect.y0, key_rect.x1, key_rect.y1
            )
            if index_key_to_embed:
                shape1.insert_textbox(
                    text_rect,
                    index_key_to_embed,
                    fontname=font_alias,
                    fontsize=10,
                    color=text_color,
                    align=0,
                )

            shape1.commit()

        # --- 2. 最終ページに コメント・書誌情報・引用QR を描画 ---
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

            # --- マージンと座標の計算 ---
            # 用紙サイズから A4/A5 を簡易判定し、適切なマージンを選択
            target_format = "A4"
            # A5幅は約420pt。500pt未満ならA5とみなす
            if paper_width < 500:
                target_format = "A5"

            margins = LAYOUT_MARGINS.get(target_format, LAYOUT_MARGINS["A4"])
            margin_left_pt = margins["left"] * CM_TO_PT
            margin_right_pt = margins["right"] * CM_TO_PT

            # 描画開始Y位置は、ヘッダー領域(key_rect)の下から + 余白
            info_rect_y_start = key_rect.y1 + 10
            info_rect_y_end = page_last.rect.height - 30
            current_y_pos = info_rect_y_start

            # X座標: 左マージン ～ 右端から右マージン分引いた位置
            x0 = margin_left_pt
            x1 = paper_width - margin_right_pt

            if sist_string_formal:
                # 書誌情報 (SIST 02)
                sist_rect = fitz.Rect(x0, current_y_pos, x1, current_y_pos + 60)
                rc_sist = shape_last.insert_textbox(
                    sist_rect,
                    f"書誌情報 (SIST 02):\n{sist_string_formal}",
                    fontname=font_alias,
                    fontsize=6,
                    align=0,
                )
                actual_sist_y1 = (
                    sist_rect.y0 + (sist_rect.height - rc_sist)
                    if rc_sist >= 0
                    else sist_rect.y1
                )
                current_y_pos = actual_sist_y1 + 10

            if sist_string_readable:
                # 書誌情報 (可読形式)
                readable_rect = fitz.Rect(x0, current_y_pos, x1, info_rect_y_end)
                rc_readable = shape_last.insert_textbox(
                    readable_rect,
                    f"書誌情報:\n{sist_string_readable}",
                    fontname=font_alias,
                    fontsize=9,
                    align=0,
                )
                actual_readable_y1 = (
                    readable_rect.y0 + (readable_rect.height - rc_readable)
                    if rc_readable >= 0
                    else readable_rect.y1
                )
                current_y_pos = actual_readable_y1 + 40
            else:
                current_y_pos = info_rect_y_start + 40

            # コメント
            comment_rect = fitz.Rect(x0, current_y_pos, x1, info_rect_y_end)
            if comment_to_embed:
                shape_last.insert_textbox(
                    comment_rect,
                    f"コメント:\n{comment_to_embed}",
                    fontname=font_alias,
                    fontsize=9,
                    align=0,
                )

            shape_last.commit()

        # 最終ページに「引用Key専用QRコード」を描画
        if cited_keys_list:
            # (page_last がまだ定義されていない場合 = コメント等が空だった場合)
            if "page_last" not in locals():
                page_last = doc.new_page(
                    pno=len(doc), width=paper_width, height=paper_height
                )

            try:
                # 1. Prepare JSON (refs only)
                qr_data_refs = {"refs": cited_keys_list}
                qr_data_str_refs = json.dumps(qr_data_refs, ensure_ascii=False)

                # 2. Generate QR image (Larger size)
                qr_refs = qrcode.QRCode(box_size=4, border=1)
                qr_refs.add_data(qr_data_str_refs)
                qr_refs.make(fit=True)
                qr_img_refs = qr_refs.make_image(fill_color="black", back_color="white")

                # 3. Get image bytes
                img_byte_arr_refs = io.BytesIO()
                qr_img_refs.save(img_byte_arr_refs, format="PNG")
                img_bytes_refs = img_byte_arr_refs.getvalue()

                # 4. Define position (Bottom-Right)
                qr_size_refs = refs_qr_size_pt  # Default: 75x75 pt (約 2.6 cm)
                margin_refs = 30  # 右下からのマージン
                qr_x_refs = page_last.rect.width - qr_size_refs - margin_refs
                qr_y_refs = page_last.rect.height - qr_size_refs - margin_refs
                qr_rect_refs = fitz.Rect(
                    qr_x_refs,
                    qr_y_refs,
                    qr_x_refs + qr_size_refs,
                    qr_y_refs + qr_size_refs,
                )

                # 5. Insert QR image
                page_last.insert_image(qr_rect_refs, stream=img_bytes_refs)
                logger.info(
                    f"引用Key専用QRコードを最終ページに埋め込みました: {Path(pdf_path_str).name}",
                    extra={"sensitive": True},
                )

            except Exception as e:
                logger.error(f"最終ページの引用Key用QRコード生成エラー: {e}")

        # 変更を上書き保存
        doc.saveIncr()

    except Exception as e:
        logger.error(
            f"Webクリップへのメタデータ埋め込み中にエラー ({pdf_path_str}): {e}",
            extra={"sensitive": True},
        )
        raise
    finally:
        if doc:
            doc.close()


def _flatten_annot_manually(page: fitz.Page, annot: fitz.Annot) -> bool:
    """
    注釈のプロパティを読み取り、PyMuPDFの描画機能でページ本体に焼き付けます。
    筆圧情報（線の強弱）は失われますが、確実に背景化できます。

    Returns:
        bool: 描画に成功し、注釈を削除した場合は True
    """
    try:
        annot_type = annot.type[0]
        rect = annot.rect

        # 色・透明度・線幅の取得 (None対策)
        colors = annot.colors if annot.colors else {}
        stroke_color = colors.get("stroke")
        fill_color = colors.get("fill")
        opacity = annot.opacity if annot.opacity is not None else 1.0

        # ボーダー情報の取得 (borderがdictでない場合やNoneの場合に対応)
        line_width = 1
        if annot.border and isinstance(annot.border, dict):
            line_width = annot.border.get("width", 1)
        elif hasattr(annot.border, "__getitem__") and len(annot.border) > 0:
            # 古いバージョン等でリストの場合
            line_width = annot.border[0]

        # --- タイプ別の描画処理 ---

        # 1. インク (手書き)
        if annot_type == fitz.PDF_ANNOT_INK:
            # vertices は「ストローク(点列)」のリスト
            if annot.vertices:
                for stroke in annot.vertices:
                    page.draw_polyline(
                        stroke,
                        color=stroke_color,
                        width=line_width,
                        stroke_opacity=opacity,
                    )

        # 2. 線
        elif annot_type == fitz.PDF_ANNOT_LINE:
            if annot.vertices and len(annot.vertices) >= 2:
                page.draw_line(
                    annot.vertices[0],
                    annot.vertices[1],
                    color=stroke_color,
                    width=line_width,
                    stroke_opacity=opacity,
                )

        # 3. 四角形 / 円
        elif annot_type == fitz.PDF_ANNOT_SQUARE:
            page.draw_rect(
                rect,
                color=stroke_color,
                fill=fill_color,
                width=line_width,
                stroke_opacity=opacity,
                fill_opacity=opacity,
            )
        elif annot_type == fitz.PDF_ANNOT_CIRCLE:
            # draw_circleは中心+半径だが、draw_ovalはRect指定で便利
            page.draw_oval(
                rect,
                color=stroke_color,
                fill=fill_color,
                width=line_width,
                stroke_opacity=opacity,
                fill_opacity=opacity,
            )

        # 4. 多角形 / 折れ線
        elif annot_type == fitz.PDF_ANNOT_POLYGON:
            if annot.vertices:
                page.draw_polygon(
                    annot.vertices,
                    color=stroke_color,
                    fill=fill_color,
                    width=line_width,
                    stroke_opacity=opacity,
                    fill_opacity=opacity,
                )
        elif annot_type == fitz.PDF_ANNOT_POLY_LINE:
            if annot.vertices:
                page.draw_polyline(
                    annot.vertices,
                    color=stroke_color,
                    width=line_width,
                    stroke_opacity=opacity,
                )

        # 5. スタンプ (画像として焼き込み)
        elif annot_type == fitz.PDF_ANNOT_STAMP:
            # 注釈の見た目をPixmapとして取得し、画像として埋め込む
            pix = annot.get_pixmap()
            page.insert_image(rect, pixmap=pix)

        # 6. フリーテキスト (テキストボックスとして焼き込み)
        elif annot_type == fitz.PDF_ANNOT_FREE_TEXT:
            text_content = annot.info.get("content", "")
            if text_content:
                # フォントサイズ取得 (0以下の場合はデフォルト設定)
                fs = annot.fontsize if annot.fontsize > 0 else 11
                # テキスト挿入 (フォントの完全再現は難しいが内容は残す)
                page.insert_textbox(
                    rect,
                    text_content,
                    color=stroke_color,
                    fontsize=fs,
                    align=fitz.TEXT_ALIGN_LEFT,
                )

        else:
            # その他の注釈はサポート外としてスキップ（削除しない）
            return False

        # 描画に成功したら、元の注釈を削除
        page.delete_annot(annot)
        return True

    except Exception as e:
        logger.warning(f"Manual flatten failed (Type {annot.type}): {e}")
        return False


def high_fidelity_flatten(
    input_path: str, output_path: str, font_path: str, flatten_ink: bool = True
) -> None:
    """
    PyMuPDFを使い、以下の処理を行います。
    1. フォーム（Widget）をテキスト化してフラット化 (常時実行)
    2. ハイライトとリンク以外の注釈（手書き等）をフラット化 (flatten_ink=Trueの場合)

    Args:
        input_path (str): 入力PDFファイルのパス。
        output_path (str): フラット化後の出力PDFファイルのパス。
        font_path (str): 埋め込むフォントファイル（.ttf, .otfなど）のパス。
        flatten_ink (bool): Trueなら手書き注釈等を背景化（筆圧消失）。Falseなら注釈のまま維持（筆圧維持）。

    Raises:
        FileNotFoundError: 指定されたフォントファイルが見つからない場合。
        Exception: PDFのオープンや保存に失敗した場合。
    """
    if not Path(font_path).is_file():
        raise FileNotFoundError(
            f"指定されたフォントファイルが見つかりません: {font_path}"
        )

    doc = None
    try:
        doc = fitz.open(input_path)
        if doc.is_encrypted:
            logger.warning(
                f"暗号化されたPDFはスキップします: {Path(input_path).name}",
                extra={"sensitive": True},
            )
            return  # 暗号化ファイルは処理せず終了

        font_name_in_pdf = "synapsen-embed-font"

        # 全ページ共通でフォント登録を試みる
        try:
            if len(doc) > 0:
                doc[0].insert_font(fontname=font_name_in_pdf, fontfile=font_path)
        except Exception:
            pass

        for page in doc:
            # --- 1. フォーム（Widget）のテキスト化 ---
            for widget in page.widgets():
                if (
                    widget.field_type
                    in (fitz.PDF_WIDGET_TYPE_TEXT, fitz.PDF_WIDGET_TYPE_COMBOBOX)
                    and widget.field_value
                ):
                    try:
                        page.insert_font(fontname=font_name_in_pdf, fontfile=font_path)
                    except Exception:
                        pass

                    page.insert_textbox(
                        widget.rect,
                        widget.field_value,
                        fontname=font_name_in_pdf,
                        fontsize=widget.text_fontsize or 10,
                        color=widget.text_color or (0, 0, 0),
                    )
                # 元のインタラクティブなウィジェットを削除
                page.delete_widget(widget)

            # --- 2. 注釈の手動フラット化 (設定依存) ---
            if flatten_ink:
                annot_list = list(page.annots())
                if annot_list:
                    for annot in annot_list:
                        # ハイライト(8) と リンク(1) は除外
                        if annot.type[0] in (
                            fitz.PDF_ANNOT_HIGHLIGHT,
                            fitz.PDF_ANNOT_LINK,
                        ):
                            continue

                        # それ以外（インク等）は手動フラット化を実行
                        _flatten_annot_manually(page, annot)

        # PDFを保存
        doc.save(output_path, garbage=4, deflate=True)

    except Exception as e:
        logger.error(
            f"フラット化処理中にエラー ({input_path}): {e}", extra={"sensitive": True}
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
    paper_height: float,
    target_format: str = "A4",
) -> None:
    """
    pypdfを使い、PDFの全ページを、指定された用紙サイズの中央にリサイズ・配置します。
    Erstellerのヘッダー・フッターと重ならないよう、マージンを考慮します。

    Args:
        input_path (str): 入力PDFファイルパス。
        output_path (str): 出力PDFファイルパス。
        paper_width (float): ターゲットの用紙幅 (ポイント単位)。
        paper_height (float): ターゲットの用紙高 (ポイント単位)。
        target_format (str): "A4" または "A5"。マージン決定に使用。
    """

    # --- 処理スキップ判定 ---
    skip_processing = False
    try:
        with fitz.open(input_path) as doc:
            keywords = doc.metadata.get("keywords", "")
            if keywords and "Synapsen:SkipNormalization" in keywords:
                skip_processing = True
    except Exception as e:
        logger.warning(f"メタデータ確認中にエラー: {e}")

    if skip_processing:
        logger.info(f"正規化スキップフラグ検出: {input_path}")
        shutil.copy2(input_path, output_path)
        return
    # -----------------------

    # マージン設定の取得
    margins = LAYOUT_MARGINS.get(target_format.upper(), LAYOUT_MARGINS["A4"])

    # cm -> pt 変換
    m_top = margins["top"] * CM_TO_PT
    m_bottom = margins["bottom"] * CM_TO_PT

    # 左右マージン緩和の設定
    min_side_margin = 0.5 * CM_TO_PT

    reader = None
    try:
        reader = PdfReader(input_path)
        writer = PdfWriter()

        # 描画可能領域（Safe Area）の計算
        drawable_height = paper_height - m_top - m_bottom
        drawable_width_relaxed = paper_width - (min_side_margin * 2)

        for content_page in reader.pages:
            # 指定された用紙サイズの白紙ページを作成
            template_page = writer.add_blank_page(
                width=paper_width, height=paper_height
            )

            original_width = float(content_page.mediabox.width)
            original_height = float(content_page.mediabox.height)

            if original_width == 0 or original_height == 0:
                logger.warning(
                    f"Skipping empty or invalid page in {input_path}",
                    extra={"sensitive": True},
                )
                continue

            # スケーリング倍率の計算
            scale_h = drawable_height / original_height
            scale_w = drawable_width_relaxed / original_width
            scale = min(scale_w, scale_h)

            # --- 配置オフセット計算 ---

            # 横方向: センタリング
            tx = (paper_width - original_width * scale) / 2

            # 縦方向: ヘッダー直下
            ty = paper_height - m_top - (original_height * scale)

            transform = (
                Transformation().scale(sx=scale, sy=scale).translate(tx=tx, ty=ty)
            )

            template_page.merge_transformed_page(content_page, transform)

        with open(output_path, "wb") as f:
            writer.write(f)

    except Exception as e:
        logger.error(
            f"正規化処理中にエラー ({input_path}): {e}", extra={"sensitive": True}
        )
        raise


# ==============================================================================
# OCR処理 (Tesseract / Ollama)
# ==============================================================================


def _run_ollama_ocr(img_data: bytes, model: str, url: str) -> str:
    """
    Ollamaライブラリを使用して画像からテキストを抽出します。
    """
    if not ollama:
        logger.error(
            "ollamaパッケージがインストールされていません。OCRを実行できません。"
        )
        return ""

    try:
        # クライアントの初期化
        # config.iniで "http://localhost:11434/api/generate" と書かれていても
        # "http://localhost:11434" だけでも動くように整形します。
        host = url.split("/api/")[0]

        client = ollama.Client(host=host)

        # 画像OCRの実行
        response = client.chat(
            model=model,
            messages=[
                {
                    "role": "user",
                    "content": "この画像に書かれている文字をすべて書き起こしてください。出力は書き起こしたテキストのみにしてください。",
                    "images": [img_data],
                }
            ],
            options={
                "temperature": 0.0,
            },
        )

        if "message" in response and "content" in response["message"]:
            return response["message"]["content"].strip()
        return ""

    except Exception as e:
        logger.error(f"Ollama OCR Error: {e}")
        return ""


def embed_ocr_text_in_pdf(
    pdf_path_str: str,
    enable_ocr: bool,
    font_path: str,
    ocr_engine: str = "tesseract",
    ollama_config: dict = None,
    lang: str = "jpn+jpn_vert",
) -> None:
    """
    PDFを解析し、既存のテキストレイヤーが存在しない場合、
    OCRを実行して結果を透明テキストレイヤーとして埋め込みます。

    Args:
        pdf_path_str (str): 処理対象のPDFファイルパス（読み書きされる）。
        enable_ocr (bool): Tesseract OCR (低速) を実行するかどうか。
        font_path (str): 埋め込む日本語フォントファイルのパス。
        ocr_engine (str): "tesseract" または "ollama"。
        lang (str): Tesseractが使用する言語。
        ollama_config (dict): Ollamaの設定辞書。

    Raises:
        Exception: Tesseract-OCRが見つからない場合。
                   その他、ファイルの上書き保存に失敗した場合。
    """
    if not enable_ocr:
        return

    doc = None
    # 一時ファイルへの保存パスを定義
    temp_output_path = pdf_path_str + "._temp_ocr.pdf"
    OCR_FONT_NAME = "synapsen_ocr_font"

    try:
        doc = fitz.open(pdf_path_str)
        if doc.is_encrypted:
            logger.info(
                f"暗号化されたPDFはスキップ: {Path(pdf_path_str).name}",
                extra={"sensitive": True},
            )
            return

        # 1. 高速なテキスト抽出を試みる
        meaningful_text_threshold = 10
        pages_processed_count = 0

        logger.info(
            f"OCR処理開始 (Engine: {ocr_engine}): {Path(pdf_path_str).name}",
            extra={"sensitive": True},
        )

        for page_num, page in enumerate(doc):
            # 既存テキストチェック
            page_text = page.get_text("text", sort=True).strip()
            if len(page_text) > meaningful_text_threshold:
                logger.debug(f"Page {page_num + 1}: 既存テキストありのためスキップ")
                continue

            pages_processed_count += 1

            # ページを画像化
            pix = page.get_pixmap(dpi=300)
            img_data = pix.tobytes("png")

            try:
                try:
                    page.insert_font(fontname=OCR_FONT_NAME, fontfile=font_path)
                except Exception:
                    pass

                # === エンジン別処理 ===
                if ocr_engine == "ollama":
                    # --- Ollama (Package) ---
                    if not ollama_config or not ollama_config.get("model"):
                        logger.warning("Ollama設定が不足しています。")
                        continue

                    logger.info(f"Ollama OCR実行中 (Page {page_num + 1})...")

                    extracted_text = _run_ollama_ocr(
                        img_data, ollama_config.get("model"), ollama_config.get("url")
                    )

                    if extracted_text:
                        pass
                    else:
                        logger.warning(
                            f"Page {page_num + 1}: Ollamaからテキストが取得できませんでした。"
                        )

                else:
                    # --- Tesseract ---
                    img = Image.open(io.BytesIO(img_data))
                    logger.info(f"Tesseract OCR実行中 (Page {page_num + 1})...")
                    tsv_data = pytesseract.image_to_data(
                        img, lang=lang, output_type=Output.STRING
                    )

                    if not tsv_data or len(tsv_data.strip()) == 0:
                        continue

                    df = pd.read_csv(
                        io.StringIO(tsv_data),
                        sep="\t",
                        quoting=csv.QUOTE_NONE,
                        on_bad_lines="skip",
                    )
                    df = df.dropna(subset=["conf", "text"])
                    df = df[df["conf"] > 30]

                    if df.empty:
                        continue

                    dpi_scale = 72 / 300
                    for _, row in df.iterrows():
                        x0, y0, w, h = (
                            row["left"],
                            row["top"],
                            row["width"],
                            row["height"],
                        )
                        rect = fitz.Rect(
                            x0 * dpi_scale,
                            y0 * dpi_scale,
                            (x0 + w) * dpi_scale,
                            (y0 + h) * dpi_scale,
                        )
                        fs = max(h * dpi_scale * 0.8, 6.0)

                        page.insert_text(
                            rect.bottom_left,
                            str(row["text"]),
                            fontname=OCR_FONT_NAME,
                            fontsize=fs,
                            render_mode=3,
                            rotate=0,
                        )

            except Exception as e_page:
                logger.error(f"Page {page_num + 1} 処理エラー: {e_page}")
                continue

        # タグ付け (Ollamaのみ)
        if ocr_engine == "ollama" and pages_processed_count > 0:
            current_metadata = doc.metadata
            keywords = current_metadata.get("keywords", "")
            ocr_tag = "Synapsen:OCR_Method=Ollama"

            if ocr_tag not in keywords:
                new_keywords = f"{keywords}; {ocr_tag}" if keywords else ocr_tag
                new_metadata = current_metadata.copy()
                new_metadata["keywords"] = new_keywords
                doc.set_metadata(new_metadata)

        # 処理が行われた場合のみ保存
        if pages_processed_count > 0:
            doc.save(
                temp_output_path,
                garbage=4,
                deflate=True,
                encryption=fitz.PDF_ENCRYPT_NONE,
            )
            doc.close()
            doc = None

            # 元ファイルへの上書き (Windowsではclose後に移動必須)
            shutil.move(temp_output_path, pdf_path_str)
            logger.info(f"OCR完了・上書き保存: {Path(pdf_path_str).name}")
        else:
            logger.info("OCR対象ページがなかったため、保存をスキップしました。")
            doc.close()
            doc = None  # ★重要

    except Exception as e:
        logger.error(f"OCR処理全体エラー ({pdf_path_str}): {e}")

        # doc.close() は finally に任せるためここでは削除するか、
        # ここで行うなら doc = None が必要ですが、削除推奨です。

        if Path(temp_output_path).is_file():
            try:
                Path(temp_output_path).unlink()
            except Exception:
                pass
        raise

    finally:
        # docが None でない（＝まだ閉じられていない）場合のみ閉じる
        if doc:
            try:
                doc.close()
            except Exception:
                pass  # 万が一の閉鎖エラーは無視

    # 12. 正常終了した場合のみ、ファイルのリネーム（上書き）
    if Path(temp_output_path).is_file():
        try:
            shutil.move(temp_output_path, pdf_path_str)
            logger.info(
                f"元ファイルに上書き完了: {Path(pdf_path_str).name}",
                extra={"sensitive": True},
            )
        except Exception as e_move:
            logger.error(
                f"PDFファイルの上書き保存に失敗 ({pdf_path_str}): {e_move}",
                extra={"sensitive": True},
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
            extra={"sensitive": True},
        )
        raise
    finally:
        if img_doc:
            img_doc.close()
        if pdf_doc:
            pdf_doc.close()


def convert_pil_image_to_pdf(pil_image: Image.Image, output_pdf_path: Path) -> None:
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
        if pil_image.mode == "RGBA":
            pil_image = pil_image.convert("RGB")  # 透過情報を除去

        pil_image.save(img_bytes_io, format="PNG")
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
    ".odt": "odt",
}


def convert_document_to_pdf(
    input_path: Path,
    output_pdf_path: Path,
    paper_size_str: str = "A4",
    pdf_margins: dict = None,
) -> None:
    """
    Pandoc (MD, TXT, DOCX等) と Playwright (HTML->PDF) を使用して ドキュメント を PDF に変換します。
    Pandoc と Playwright (chromium) がインストールされている必要があります。
    変換前に <details> を <details open> に置換します。
    又、CSSリセットを適用し、Playwright生成時の余白を排除します。

    Args:
        input_path (Path): 入力ファイルのパス。
        output_pdf_path (Path): 出力先PDFファイルのパス。
        paper_size_str (str): "A4" または "A5"。
        pdf_margins (dict, optional): Playwrightのpage.pdfに渡すマージン設定。
                                      Noneの場合はデフォルト(上下左右0)を使用。
    """

    # 一時ファイル用のパスを定義
    temp_modified_content_path = output_pdf_path.with_suffix(".temp.modified")
    temp_html_path = output_pdf_path.with_suffix(".temp.html")

    file_suffix = input_path.suffix.lower()

    # デフォルトマージンの設定 (Playwright用)
    if pdf_margins is None:
        pdf_margins = {"top": "0", "bottom": "0", "left": "0", "right": "0"}

    # --- ステップ 1: 前処理 (フォーマットごとに行う) ---
    try:
        with open(input_path, "r", encoding="utf-8") as f:
            content = f.read()

        # Markdownの場合のみ <details> を置換
        if file_suffix == ".md":
            modified_content = re.sub(
                r"<details(?![^>]*\bopen\b)",
                "<details open",
                content,
                flags=re.IGNORECASE,
            )

        # テキストファイルの場合
        elif file_suffix == ".txt":
            import html

            escaped = html.escape(content)
            # preタグで囲むことで改行とフォントを維持
            modified_content = (
                "<pre style='white-space: pre-wrap; font-family: monospace;'>"
                f"{escaped}</pre>"
            )
        else:
            modified_content = content

        # 前処理が不要な形式（docxなど）と、処理済みの内容を一時ファイルに書き出す
        if file_suffix in [".docx", ".rtf", ".odt"]:
            # バイナリファイルをコピー
            shutil.copy2(input_path, temp_modified_content_path)
        else:
            # テキストベースのファイルを書き出し
            with open(temp_modified_content_path, "w", encoding="utf-8") as f:
                f.write(modified_content)

    except Exception as e:
        # DOCXなどは 'utf-8' で読めないため、バイナリとして扱う
        if file_suffix in PANDOC_INPUT_FORMATS and file_suffix not in [".md", ".txt"]:
            try:
                shutil.copy2(input_path, temp_modified_content_path)
            except Exception as copy_e:
                raise Exception(
                    f"ドキュメントの前処理（コピー）に失敗しました: {copy_e}"
                )
        else:
            raise Exception(f"ドキュメントの前処理（読み込み）に失敗しました: {e}")

    # --- ステップ 2: Pandoc で HTML (一時ファイル) に変換 ---
    # マッピング辞書から入力フォーマットを取得
    input_format = PANDOC_INPUT_FORMATS.get(file_suffix, "gfm")  # 不明な場合はgfm扱い

    # .txtファイルの場合、自分で<pre>タグでHTML化しているので、入力形式を'html'として扱う
    if file_suffix == ".txt":
        input_format = "html"

    pandoc_cmd = [
        "pandoc",
        "--from",
        input_format,
        str(temp_modified_content_path),
        "-s",
        "--embed-resources",
        "--mathml",
        "--to",
        "html5",
        "-o",
        str(temp_html_path),
    ]
    logger.info(
        f"Pandoc (MD->HTML) 実行: {' '.join(pandoc_cmd)}", extra={"sensitive": True}
    )

    try:
        subprocess.run(
            pandoc_cmd,
            capture_output=True,
            text=True,
            encoding="utf-8",
            errors="ignore",
            check=True,
        )
    except FileNotFoundError:
        # finallyブロックで一時MDファイルが削除されるよう、エラーを再送出
        raise Exception(
            "Pandoc が見つかりません。\n"
            + "Markdown連携には Pandoc のインストールとPATH設定が必要です。"
        )
    except subprocess.CalledProcessError as e:
        error_details = f"STDOUT:\n{e.stdout}\n\nSTDERR:\n{e.stderr}"
        # finallyブロックで一時MDファイルが削除されるよう、エラーを再送出
        raise Exception(
            f"PandocでのHTML変換に失敗しました (ReturnCode {e.returncode}):\n"
            f"{error_details}"
        )

    # --- 生成されたHTMLにCSSを注入 ---
    reset_css = """
    <style>
        /* 全要素のボックスサイズ計算を統一 */
        * {
            box-sizing: border-box;
        }

        /* html, body の幅を100%にし、余白を除去 */
        html, body {
            margin: 0 !important;
            padding: 0 !important;
            width: 100% !important;
            height: 100% !important;
            background-color: white !important;
        }

        /* Pandoc等の幅制限解除 & 日本語改行ルールの適用 */
        body, main, article, div, .markdown-body, .main-content, p, li, dd, dt, th, td {
            max-width: none !important;
            margin-left: 0 !important;
            margin-right: 0 !important;
            width: 100% !important;

            /* ★変更点: 日本語の「原稿用紙」的な挙動にする設定 */
            word-break: break-all !important; /* 単語の途中でも文字数に合わせて改行 */
            line-break: strict !important;    /* 句読点や括弧が行頭に来ないよう厳格に処理 */
            text-align: justify !important;   /* 両端揃え（行末を揃える） */
        }

        /* 本文に適度なパディングを設定 */
        body {
            padding: 1.5rem !important;
        }

        /* コードブロック等は折り返しつつ、英単語の途中では切らない方が読みやすい場合も多いが
           今回は全体の統一感を優先して break-all に設定（必要に応じて break-word に戻せます） */
        pre, code, pre code {
            white-space: pre-wrap !important;
            word-break: break-all !important;
            max-width: 100% !important;
            margin: 0 !important;
        }

        /* 画像調整 */
        img {
            max-width: 100% !important;
            height: auto !important;
        }
    </style>
    """

    try:
        with open(temp_html_path, "r", encoding="utf-8") as f:
            html_content = f.read()

        # </head> の直前に CSS を挿入
        if "</head>" in html_content:
            html_content = html_content.replace("</head>", f"{reset_css}\n</head>")
        else:
            # headがない場合は先頭に追加
            html_content = reset_css + html_content

        with open(temp_html_path, "w", encoding="utf-8") as f:
            f.write(html_content)

    except Exception as e:
        logger.warning(f"CSS注入に失敗しましたが続行します: {e}")

    # --- ステップ 3: Playwright で HTML を PDF に変換 ---
    playwright_paper_format = paper_size_str.upper()

    pw_instance = None
    browser = None
    page = None
    try:
        logger.info(
            f"Playwright (HTML->PDF) 実行: {temp_html_path.name}",
            extra={"sensitive": True},
        )
        pw_instance = sync_playwright().start()
        browser = pw_instance.chromium.launch()
        page = browser.new_page()

        # HTMLファイルを開く
        page.goto(temp_html_path.as_uri(), wait_until="networkidle")

        # PDF保存
        page.pdf(
            path=str(output_pdf_path),
            format=playwright_paper_format,
            print_background=True,
            margin=pdf_margins,
        )
        logger.info(
            f"Playwright PDF変換完了: {output_pdf_path.name}", extra={"sensitive": True}
        )

    except PlaywrightError as e:
        # finallyブロックで一時MDファイルが削除されるよう、エラーを再送出
        raise Exception(
            "Playwright (Chromium) でのHTML->PDF変換に失敗しました。\n"
            + "Install_Poetry.bat を実行して Playwright が正しくインストールされているか確認してください。\n"
            + f"エラー: {e}"
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

        # 一時ファイルの削除
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
                    extra={"sensitive": True},
                )

import re
from pathlib import Path
from pypdf import PdfReader
import fitz  # PyMuPDF

# --- 追加インポート ---
from PIL import Image
import io
try:
    from pyzbar.pyzbar import decode
except ImportError:
    decode = None  # ライブラリがない場合のフォールバック用

import logging
logger = logging.getLogger(__name__)


# ==============================================================================
# PDF情報取得関数
# ==============================================================================
def get_note_info(pdf_path: Path, key_rect: tuple):
    """
    単一のPDFファイルを解析し、ファイル名や内容から情報を抽出する。
    優先順位: QRコード -> 指定座標のテキスト (key_rect)
    """
    try:
        commonplace_key = ""
        full_text = ""  # full_text は空で初期化 (抽出しない)
        doc = None
        try:
            # PyMuPDF を使って Index Key のみ取得
            doc = fitz.open(pdf_path)
            if len(doc) > 0:
                page = doc[0]

                # --- 1. QRコードからの読み取り (優先) ---
                if decode is not None:
                    try:
                        # 処理高速化のため DPI=150 程度でレンダリング
                        pix = page.get_pixmap(dpi=150)
                        img_data = pix.tobytes("png")
                        pil_image = Image.open(io.BytesIO(img_data))

                        decoded_objects = decode(pil_image)
                        for obj in decoded_objects:
                            if obj.type == 'QRCODE':
                                qr_text = obj.data.decode('utf-8').strip()
                                if qr_text:
                                    commonplace_key = qr_text
                                    logger.debug(
                                        f"QR検出: {commonplace_key} "
                                        f"({pdf_path.name})"
                                    )
                                    break
                    except Exception as e:
                        # 画像処理エラー等はログに出して無視し、テキスト抽出へ進む
                        logger.debug(f"QR読み取り失敗 ({pdf_path.name}): {e}")
                else:
                    # pyzbar がインストールされていない場合は初回のみ警告ログを出す等の処理も可能
                    pass

                # --- 2. 指定座標からのテキスト抽出 (フォールバック) ---
                # QRコードでキーが見つからなかった場合のみ実行
                if not commonplace_key and key_rect and len(key_rect) == 4:
                    commonplace_key = page.get_textbox(key_rect).strip()

        except Exception as e:
            logger.error(f"PyMuPDFでのIndex Key抽出エラー ({pdf_path.name}): {e}")
        finally:
            if doc:
                doc.close()  # 確実に閉じる

        # PyPDF でページ数を取得
        page_count = len(PdfReader(pdf_path).pages)
        filepath = str(pdf_path)

        match = re.match(
            r"(\d{8})_(?:(\d{4,6})_)?(.+)\.pdf",
            pdf_path.name,
            re.IGNORECASE)

        auto_generated_key = ""
        if match:
            date_str, time_val, _ = match.groups()
            # YYYYMMDDhhmmss形式のユニークIDを生成
            # timeがファイル名にない場合は '000000' で補完
            time_str = time_val.ljust(6, '0') if time_val else "999999"
            key_time = time_str if time_str != "999999" else "000000"
            auto_generated_key = date_str + key_time

        common_data = {
            "pages": page_count,
            "tags": [],
            "key": auto_generated_key,
            "memo": "",
            "commonplace_key": commonplace_key,
            "filepath": filepath,
            "full_text": full_text  # <-- ここでは空文字が設定される
        }
        if not match:
            return {
                "date": "日付不明",
                "time": "999999",
                "title": pdf_path.stem,
                **common_data,
                "is_warning": True
                }

        date_str, time_val, title = match.groups()
        time_str = time_val.ljust(6, '0') if time_val else "999999"
        return {
            "date": date_str,
            "time": time_str,
            "title": title,
            **common_data,
            "is_warning": False
        }

    except Exception as e:
        logger.error(f"PDF情報取得エラー ({pdf_path.name}): {e}")
        return {
            "date": "読み込み失敗",
            "time": "999999",
            "title": pdf_path.name,
            "pages": 0,
            "tags": [],
            "key": "",
            "memo": "",
            "commonplace_key": "",
            "filepath": str(pdf_path),
            "full_text": "",
            "is_warning": True
        }


# ==============================================================================
# full_text 抽出ヘルパー関数
# ==============================================================================
def get_full_text(pdf_path: Path) -> str:
    """
    指定されたPDFファイルから埋め込みテキスト（full_text）のみを抽出する。
    get_note_info のテキスト抽出部分と同一のロジック。

    Args:
        pdf_path (Path): 処理対象のPDFパス。

    Returns:
        str: 抽出されたテキスト。
    """
    full_text = ""
    doc = None
    try:
        doc = fitz.open(pdf_path)
        # 全ページのテキストを抽出 (埋め込まれたOCRテキストを含む)
        for page in doc:
            full_text += page.get_text("text", sort=True) + "\n"
        return full_text.strip()
    except Exception as e:
        logger.error(f"PyMuPDFでのテキスト抽出エラー ({pdf_path.name}): {e}")
        return ""  # エラー時は空文字を返す
    finally:
        if doc:
            doc.close()  # 確実に閉じる

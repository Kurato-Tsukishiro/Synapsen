import re
from pathlib import Path
from pypdf import PdfReader
import fitz  # PyMuPDF
import unicodedata
import json

# --- 追加インポート ---
from PIL import Image
import io
try:
    from pyzbar.pyzbar import decode
except ImportError:
    decode = None  # ライブラリがない場合のフォールバック用

import logging
logger = logging.getLogger(__name__)


def _normalize_key_text(raw_text: str) -> str:
    """
    Index Keyとして読み取ったテキストから不要文字を除去し、正規化する。
    """
    if not raw_text:
        return ""

    # 1. 日本語(ひらがな,カタカナ,漢字), 英数字(\w), 記号(/・),
    #    半角スペース以外のすべての文字(制御文字, ゼロ幅スペース等)を除去
    #    (\w は英数字とアンダースコアにマッチします)
    pattern_to_remove = (
        r'[^\w\u3040-\u30FF\u3400-\u4DBF\u4E00-\u9FFF/\s・]'
    )
    cleaned_text = re.sub(
        pattern_to_remove,
        '',
        raw_text
    )

    # 2. 改行や連続する空白を単一のスペースに正規化
    normalized_text = (
        " ".join(cleaned_text.split())
        .strip()
    )

    # 3. Unicode正規化 (NFKC)
    #    互換文字 (例: '識' U+F9BC) を 標準文字 (例: '識' U+8B58) に変換
    final_text = unicodedata.normalize(
        'NFKC', normalized_text)

    return final_text


# ==============================================================================
# PDF情報取得関数
# ==============================================================================
def get_note_info(pdf_path: Path, key_rect: tuple):
    """
    単一のPDFファイルを解析し、ファイル名や内容から情報を抽出する。
    優先順位: JSON QRコード -> key_rect テキスト -> ファイル名
    """
    try:
        # --- 1. まずファイル名を解析 (フォールバック用) ---
        match = re.match(
            r"(\d{8})_(?:(\d{4,6})_)?(.+)\.pdf",
            pdf_path.name,
            re.IGNORECASE)

        if match:
            date_str, time_val, title = match.groups()
            time_str = time_val.ljust(6, '0') if time_val else "999999"
            key_time = time_str if time_str != "999999" else "000000"
            auto_generated_key = date_str + key_time
            is_warning = False
        else:
            date_str = "日付不明"
            time_str = "999999"
            title = pdf_path.stem
            auto_generated_key = ""
            is_warning = True

        # --- 2. PDFからメタデータを抽出 (QR/テキスト) ---
        commonplace_key = ""
        doc = None
        try:
            doc = fitz.open(pdf_path)
            if len(doc) > 0:
                page = doc[0]
                qr_found = False

                # --- 2A. QRコードからの読み取り (優先) ---
                # Normalisierer のクリップ機能で埋め込まれたQRコードを想定
                # clipperのOCRが、埋め込んだIndex Keyと干渉する事があるため、QRコードを付与している
                if decode is not None:
                    try:
                        pix = page.get_pixmap(dpi=200)
                        img_data = pix.tobytes("png")
                        pil_image = Image.open(io.BytesIO(img_data))

                        decoded_objects = decode(pil_image)

                        if not decoded_objects:
                            logger.warning(
                                f"QR: pyzbarは起動しましたが、QRを検出できませんでした "
                                f"({pdf_path.name})",
                                extra={'sensitive': True}
                            )

                        for obj in decoded_objects:
                            if obj.type == 'QRCODE':
                                qr_text = obj.data.decode('utf-8').strip()
                                if qr_text:
                                    try:
                                        # 【JSONパース試行】
                                        qr_data = json.loads(qr_text)
                                        # cpk (IndexKey) の取得
                                        if "cpk" in qr_data:
                                            commonplace_key = (
                                                unicodedata.normalize(
                                                    'NFKC',
                                                    qr_data.get("cpk", "")
                                                )
                                            )

                                        # key (ユニークID) の取得
                                        if "key" in qr_data:
                                            auto_generated_key = qr_data["key"]
                                            # key がQRから取得できたら、
                                            # ファイル名が正規表現にマッチしなくても警告解除
                                            is_warning = False

                                        qr_found = True
                                        logger.info(
                                            f"QR(JSON)読み取り成功: "
                                            f"cpk={commonplace_key}, "
                                            f"key={auto_generated_key} "
                                            f"({pdf_path.name})",
                                            extra={'sensitive': True}
                                        )

                                    except json.JSONDecodeError:
                                        # 【フォールバック: 文字列のみ】
                                        commonplace_key = (
                                            unicodedata.normalize(
                                                'NFKC',
                                                qr_text
                                            )
                                        )
                                        qr_found = True
                                        logger.info(
                                            "QR(非JSON)読み取り成功: "
                                            f"{commonplace_key} "
                                            f"({pdf_path.name})",
                                            extra={'sensitive': True}
                                        )

                                    break
                    except Exception as e:
                        logger.warning(
                            f"QR読み取り失敗 ({pdf_path.name}): {e}",
                            extra={'sensitive': True}
                        )
                else:
                    logger.warning(
                        "QRデバッグ: pyzbarがインポートされていない (decode is None)"
                    )

                # --- 2B. 指定座標からのテキスト抽出 (基本処理) ---
                if not qr_found and key_rect and len(key_rect) == 4:
                    logger.warning(
                        f"QRデバッグ: QRが見つからなかったため、key_rectを検索します "
                        f"({pdf_path.name})",
                        extra={'sensitive': True}
                    )
                    raw_text = page.get_textbox(key_rect)
                    commonplace_key = _normalize_key_text(raw_text)

        except Exception as e:
            logger.error(
                f"PyMuPDFでのIndex Key抽出エラー ({pdf_path.name}): {e}",
                extra={'sensitive': True}
            )
        finally:
            if doc:
                doc.close()

        # --- 3. ページ数取得と辞書作成 ---
        page_count = len(PdfReader(pdf_path).pages)
        filepath = str(pdf_path)
        return {
            "date": date_str,
            "time": time_str,
            "title": title,
            "pages": page_count,
            "tags": [],
            "key": auto_generated_key,
            "memo": "",
            "commonplace_key": commonplace_key,
            "filepath": filepath,
            "full_text": "",
            "is_warning": is_warning
        }

    except Exception as e:
        logger.error(
            f"PDF情報取得エラー ({pdf_path.name}): {e}",
            extra={'sensitive': True}
        )
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
        # 全ページのテキストを抽出
        for page in doc:
            full_text += page.get_text("text", sort=True) + "\n"

        # 抽出したfull_text全体も正規化する
        if full_text:
            normalized_text = unicodedata.normalize('NFKC', full_text)
            return normalized_text.strip()
        else:
            return ""

    except Exception as e:
        logger.error(
            f"PyMuPDFでのテキスト抽出エラー ({pdf_path.name}): {e}",
            extra={'sensitive': True}
        )
        return ""
    finally:
        if doc:
            doc.close()

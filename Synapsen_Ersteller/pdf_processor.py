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
    pattern_to_remove = r"[^\w\u3040-\u30FF\u3400-\u4DBF\u4E00-\u9FFF/\s・]"
    cleaned_text = re.sub(pattern_to_remove, "", raw_text)

    # 2. 改行や連続する空白を単一のスペースに正規化
    normalized_text = " ".join(cleaned_text.split()).strip()

    # 3. Unicode正規化 (NFKC)
    #    互換文字 (例: '識' U+F9BC) を 標準文字 (例: '識' U+8B58) に変換
    final_text = unicodedata.normalize("NFKC", normalized_text)

    return final_text


# ==============================================================================
# PDF情報取得関数
# ==============================================================================
def get_note_info(pdf_path: Path, key_rect: tuple):
    """
    単一のPDFファイルを解析し、メタデータ、QRコード、またはテキストから情報を抽出する。
    優先順位: PDFメタデータ(JSON) -> QRコード -> key_rect テキスト -> ファイル名
    """
    try:
        # --- 1. ファイル名解析 ---
        match = re.match(
            r"(\d{8})_(?:(\d{4,6})_)?(.*)\.pdf", pdf_path.name, re.IGNORECASE
        )

        if match:
            date_str, time_val, title = match.groups()
            time_str = time_val.ljust(6, "0") if time_val else "999999"
            key_time = time_str if time_str != "999999" else "000000"
            auto_generated_key = date_str + key_time
            is_warning = False

            # タイトルが空、または記号のみの場合は空文字にする
            if not title or not title.strip():
                title = ""
        else:
            date_str = "日付不明"
            time_str = "999999"
            title = pdf_path.stem

            # ".pdf" のようなドット始まりや空文字はタイトルとみなさない
            if title.startswith(".") or not title.strip():
                title = ""

            auto_generated_key = ""
            is_warning = True

        # 変数初期化
        commonplace_key = ""
        memo_from_refs = ""  # 引用リンクなど
        summary_content = ""  # コメント/概要
        auto_detected_tags = []  # 自動検出されたタグを格納するリスト

        doc = None

        # ハイブリッド処理のステータス管理
        # (メタデータが見つかれば、後の重い処理をスキップするため)
        data_found_priority = False

        try:
            doc = fitz.open(pdf_path)
            if len(doc) > 0:

                # ============================================================
                # Phase 1: PDFメタデータ(Subject)からの高速読み取り
                # ============================================================
                metadata = doc.metadata

                # --- Keywords からの自動タグ付けロジック ---
                keywords = metadata.get("keywords", "") or ""

                # Canvas (全体) -> STypes_Canvas
                if "Synapsen:Whiteboard" in keywords:
                    if "STypes_Canvas" not in auto_detected_tags:
                        auto_detected_tags.append("STypes_Canvas")

                # Sticky (付箋) -> STypes_Sticky
                if "Synapsen:Sticky" in keywords:
                    if "STypes_Sticky" not in auto_detected_tags:
                        auto_detected_tags.append("STypes_Sticky")

                # WebClip -> SType_WebClip AND ZTypes_Source
                if "Synapsen:WebClip" in keywords:
                    # 出自を示すタグ
                    if "SType_WebClip" not in auto_detected_tags:
                        auto_detected_tags.append("SType_WebClip")
                    # 役割を示すタグ
                    if "ZTypes_Source" not in auto_detected_tags:
                        auto_detected_tags.append("ZTypes_Source")

                # Source (書誌情報あり等) -> ZTypes_Source
                if "Synapsen:Source" in keywords:
                    if "ZTypes_Source" not in auto_detected_tags:
                        auto_detected_tags.append("ZTypes_Source")
                # ----------------------------------------------------

                subject = metadata.get("subject", "") or ""

                # 隠しJSONを探す
                json_match = re.search(
                    r"<synapsen>(.*?)</synapsen>", subject, re.DOTALL
                )

                if json_match:
                    try:
                        json_str = json_match.group(1)
                        meta_data = json.loads(json_str)

                        # Index Key
                        if "cpk" in meta_data and meta_data["cpk"]:
                            commonplace_key = unicodedata.normalize(
                                "NFKC", meta_data["cpk"]
                            )

                        # Unique Key (ファイル名と異なる場合、こちらを優先)
                        if "key" in meta_data and meta_data["key"]:
                            auto_generated_key = meta_data["key"]
                            is_warning = False

                        # 引用 (Refs) -> メモ欄へ
                        if "refs" in meta_data and isinstance(meta_data["refs"], list):
                            links = [f"[[{r}]]" for r in meta_data["refs"]]
                            if links:
                                memo_from_refs = "\n".join(links) + "\n"

                        # コメント -> Summary欄へ
                        if "comment" in meta_data and meta_data["comment"]:
                            summary_content = meta_data["comment"]

                        data_found_priority = True
                        logger.info(
                            f"Metadata(JSON) Hit: {pdf_path.name}",
                            extra={"sensitive": True},
                        )

                    except Exception as e:
                        logger.warning(f"Metadata parse error ({pdf_path.name}): {e}")

                # ============================================================
                # Phase 2: QRコード / テキスト解析
                # ============================================================
                if not data_found_priority:
                    qr_found = False

                    # 2A. QRコード (Page 1)
                    if decode is not None:
                        try:
                            page1 = doc[0]
                            pix1 = page1.get_pixmap(dpi=200)
                            img_data1 = pix1.tobytes("png")
                            pil_image1 = Image.open(io.BytesIO(img_data1))
                            decoded_objects1 = decode(pil_image1)

                            for obj in decoded_objects1:
                                if obj.type == "QRCODE":
                                    qr_text = obj.data.decode("utf-8").strip()
                                    if qr_text:
                                        try:
                                            # JSON形式
                                            qr_data = json.loads(qr_text)
                                            if "cpk" in qr_data:
                                                commonplace_key = unicodedata.normalize(
                                                    "NFKC", qr_data.get("cpk", "")
                                                )
                                            if "key" in qr_data:
                                                auto_generated_key = qr_data["key"]
                                                is_warning = False
                                            qr_found = True
                                            logger.info(
                                                f"QR(Page1) Hit: {pdf_path.name}",
                                                extra={"sensitive": True},
                                            )
                                        except json.JSONDecodeError:
                                            # 旧形式(文字列のみ)
                                            commonplace_key = unicodedata.normalize(
                                                "NFKC", qr_text
                                            )
                                            qr_found = True
                                    break
                        except Exception as e:
                            logger.warning(f"QR scan error (Page 1): {e}")

                    # 2B. QRコード (Last Page)
                    if len(doc) > 1 and decode is not None:
                        try:
                            pageLast = doc[-1]
                            pixLast = pageLast.get_pixmap(dpi=200)
                            img_dataLast = pixLast.tobytes("png")
                            pil_imageLast = Image.open(io.BytesIO(img_dataLast))
                            decoded_objectsLast = decode(pil_imageLast)

                            for obj in decoded_objectsLast:
                                if obj.type == "QRCODE":
                                    qr_text_refs = obj.data.decode("utf-8").strip()
                                    if qr_text_refs:
                                        try:
                                            qr_data_refs = json.loads(qr_text_refs)
                                            if "refs" in qr_data_refs and isinstance(
                                                qr_data_refs["refs"], list
                                            ):
                                                links = [
                                                    f"[[{rk}]]"
                                                    for rk in qr_data_refs["refs"]
                                                ]
                                                if links:
                                                    memo_from_refs = (
                                                        "\n".join(links) + "\n"
                                                    )
                                                    logger.info(
                                                        "QR(LastPage) Hit: "
                                                        f"{len(links)} refs",
                                                    )
                                        except json.JSONDecodeError:
                                            pass
                                        break
                        except Exception as e:
                            logger.warning(f"QR scan error (Last Page): {e}")

                    # 2C. テキスト抽出
                    if not qr_found and key_rect and len(key_rect) == 4:
                        try:
                            raw_text = doc[0].get_textbox(key_rect)
                            if raw_text.strip():
                                commonplace_key = _normalize_key_text(raw_text)
                                logger.info(
                                    f"Text Rect Hit: {commonplace_key}",
                                    extra={"sensitive": True},
                                )
                        except Exception:
                            pass

        except Exception as e:
            logger.error(
                f"PDF parse error ({pdf_path.name}): {e}", extra={"sensitive": True}
            )
        finally:
            if doc:
                doc.close()

        # --- タイトルからの MOC 自動検出 (STypes_MOC) ---
        if "MOC" in title.upper() or "Map of Content" in title:
            if "STypes_MOC" not in auto_detected_tags:
                auto_detected_tags.append("STypes_MOC")

        # ページ数取得 (pypdf)
        try:
            page_count = len(PdfReader(pdf_path).pages)
        except Exception:
            page_count = 0

        return {
            "date": date_str,
            "time": time_str,
            "title": title,
            "pages": page_count,
            "tags": auto_detected_tags,
            "key": auto_generated_key,
            "memo": memo_from_refs,
            "commonplace_key": commonplace_key,
            "filepath": str(pdf_path),
            "full_text": "",
            "summary": summary_content,
            "is_warning": is_warning,
        }

    except Exception as e:
        logger.error(
            f"Fatal error processing {pdf_path.name}: {e}", extra={"sensitive": True}
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
            "summary": "",
            "is_warning": True,
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
            return clean_ocr_text(full_text)
        else:
            return ""

    except Exception as e:
        logger.error(
            f"PyMuPDFでのテキスト抽出エラー ({pdf_path.name}): {e}",
            extra={"sensitive": True},
        )
        return ""
    finally:
        if doc:
            doc.close()


# ==============================================================================
# テキスト補正関数
# ==============================================================================
def clean_ocr_text(text: str) -> str:
    """
    OCR由来のテキストに含まれる不要な空白やノイズを除去し、読みやすく整形する。
    """
    if not text:
        return ""

    # 0. 改行コードを LF (\n) に統一
    text = text.replace("\r\n", "\n").replace("\r", "\n")

    # --- ヘッダー・フッター除去処理 ---
    # 構成: icon日付 Titel(x/y)z  (zは改行されている場合や無い場合も考慮)
    # 対応パターン例:
    # 1. 1行: icon日付 Titel(x/y)z
    # 2. 2行: icon日付 Titel\n(x/y)z
    # 3. 3行: icon日付 Titel\n(x/y)\nz

    text = re.sub(
        (
            r"^.*?\d{4}/\d{2}/\d{2}"    # 行頭から日付まで
            r".*?"                      # タイトル (改行の手前まで)
            r"(?:\n|\s+)"               # 区切り (改行 または 空白)
            r"\(\s*\d+\s*/\s*\d+\s*\)"  # (元のページ/元の総ページ)
            r"(?:(?:\n|\s+)\d+)?"       # オプションのフッターz (改行/空白 + 数字)
            r".*?$\n?"                  # 残りの余分な文字と、削除後の詰め用改行
        ),
        "",
        text,
        flags=re.MULTILINE,
    )
    # ---------------------------------------------

    # 1. Unicode正規化 (NFKC)
    text = unicodedata.normalize("NFKC", text)

    # 2. 前後の空白削除
    text = text.strip()

    # 3. 日本語文字間の不要な半角スペースを除去
    text = re.sub(r"([ぁ-んァ-ン一-龥])\s+([ぁ-んァ-ン一-龥])", r"\1\2", text)

    # 4. 文中の不要な改行を削除して連結
    #    削除条件:
    #    1. 直前が句読点等ではない (?<=[^。！？…])
    #    2. 直後に改行が続かない   (?!\n)
    #    3. 直後がリスト記号ではない (?:- |・|\* |\d+\. )
    #       (ハイフン、中黒、アスタリスク、数字+ドット+空白)
    text = re.sub(r"(?<=[^。！？…])\n(?!\n)(?!\s*(?:- |・|\* |\d+\. ))", "", text)

    # 5. 文末記号の直後に改行を挿入
    text = re.sub(r"([。？！」]|(?<!\d)\.+(?!\.))\n*", r"\1\n", text)

    # 6. 連続する空白やタブを単一スペースに置換
    text = re.sub(r"[ \t]+", " ", text)

    # 7. 連続する改行を2つまでに制限 (無駄に広すぎる行間を詰める)
    text = re.sub(r"\n{3,}", "\n\n", text)

    return text

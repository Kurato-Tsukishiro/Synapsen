import re
from pathlib import Path
from pypdf import PdfReader
import fitz  # PyMuPDF


# ==============================================================================
# PDF情報取得関数
# ==============================================================================
def get_note_info(pdf_path: Path, key_rect: tuple):
    """
    単一のPDFファイルを解析し、ファイル名や内容から情報を抽出する。
    """
    try:
        commonplace_key = ""
        full_text = ""
        doc = None
        try:
            doc = fitz.open(pdf_path)
            if len(doc) > 0:
                page = doc[0]
                if key_rect and len(key_rect) == 4:
                    commonplace_key = page.get_textbox(key_rect).strip()

            # 全ページのテキストを抽出 (埋め込まれたOCRテキストを含む)
            for page in doc:
                full_text += page.get_text("text", sort=True) + "\n"
            full_text = full_text.strip()

        except Exception as e:
            print(f"PyMuPDFでのテキスト抽出エラー ({pdf_path.name}): {e}")
        finally:
            if doc:
                doc.close()  # 確実に閉じる

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
            "full_text": full_text
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
        print(f"PDF情報取得エラー ({pdf_path.name}): {e}")
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
        print(f"PyMuPDFでのテキスト抽出エラー ({pdf_path.name}): {e}")
        return ""  # エラー時は空文字を返す
    finally:
        if doc:
            doc.close()  # 確実に閉じる

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
        try:
            doc = fitz.open(pdf_path)
            if len(doc) > 0:
                page = doc[0]
                if key_rect and len(key_rect) == 4:
                    commonplace_key = page.get_textbox(key_rect).strip()
            doc.close()
        except Exception as e:
            print(f"PyMuPDFでのテキスト抽出エラー ({pdf_path.name}): {e}")

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
            "filepath": filepath
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
            "is_warning": True
        }

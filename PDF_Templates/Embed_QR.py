import fitz  # PyMuPDF
import sys
import os
import json
import configparser
import qrcode
import io
import argparse
from pathlib import Path

# --- デフォルト設定 (config.iniが見つからない場合のフォールバック) ---
DEFAULT_FONT_PATH = r"C:\Windows\Fonts\msgothic.ttc"
# Webclipのデフォルト設定 (pdf_utils.py準拠)
DEFAULT_KEY_RECT = (0, 13, 391, 73)


def load_synapsen_config(base_dir):
    """
    スクリプトの場所を基準に config.ini を探して読み込む関数
    """
    # 探すパスの候補 (同階層、1つ上、2つ上)
    candidates = [
        base_dir / "config.ini",
        base_dir / "../config.ini",
        base_dir / "../../config.ini",
    ]

    config = configparser.ConfigParser(interpolation=None)
    found_path = None

    for p in candidates:
        if p.resolve().exists():
            found_path = p.resolve()
            break

    if not found_path:
        return None, None, None, None, None

    config.read(found_path, encoding="utf-8")

    # Index Key
    keys = []
    if config.has_section("CommonplaceKeys"):
        opts = config.get("CommonplaceKeys", "options", fallback="")
        keys = [k.strip() for k in opts.split(",") if k.strip()]

    # Colors
    colors = {}
    if config.has_section("KeyColors"):
        colors = {k.lower(): v for k, v in config.items("KeyColors")}

    # Icons
    icons = {}
    if config.has_section("KeyIcons"):
        icons = {k.lower(): v for k, v in config.items("KeyIcons")}

    # Font Path
    font_path = None
    if config.has_section("Paths"):
        fp = config.get("Paths", "font_path", fallback="")
        if fp:
            # ここで環境変数を展開します
            font_path = os.path.expandvars(fp)

    # Key Rect (Webclipと同じ位置設定)
    key_rect = DEFAULT_KEY_RECT
    if config.has_section("Extraction"):
        rect_str = config.get("Extraction", "key_rect", fallback="")
        if rect_str:
            try:
                key_rect = tuple(map(float, rect_str.split(",")))
            except ValueError:
                pass

    return keys, colors, icons, font_path, key_rect


def hex_to_rgb(hex_code):
    """#RRGGBB 形式を (R, G, B) のタプル (0.0~1.0) に変換"""
    if not hex_code.startswith("#"):
        return (0, 0, 0)
    hex_code = hex_code.lstrip("#")
    return tuple(int(hex_code[i : i + 2], 16) / 255.0 for i in (0, 2, 4))


def process_pdf(input_pdf, output_dir, keys, colors, icons, font_path, key_rect):
    """指定されたPDFに各KeyのQRコードを埋め込んで保存する"""

    # フォントの決定 (config優先 -> なければデフォルト)
    use_font_path = (
        font_path if (font_path and os.path.exists(font_path)) else DEFAULT_FONT_PATH
    )
    font_name = "embed_font"

    print(f"入力ファイル: {input_pdf.name}")
    print(f"出力フォルダ: {output_dir}")

    # PDFを一回開いてメタデータなどを確認
    try:
        src_doc = fitz.open(input_pdf)
        if len(src_doc) < 1:
            return
        src_doc.close()
    except Exception as e:
        print(f"エラー: PDFを開けませんでした: {e}")
        return

    # Webclip (pdf_utils.py) の配置ロジック
    # -------------------------------------------------
    qr_size = 35  # 固定値 (pdf_utils.py準拠)
    qr_margin = 5  # 固定値

    # QRコードの座標計算
    # key_rect = (x0, y0, x1, y1)
    qr_x = key_rect[0] + qr_margin
    qr_y = key_rect[1]
    qr_rect = fitz.Rect(qr_x, qr_y, qr_x + qr_size, qr_y + qr_size)

    # テキストの座標計算 (QRの右隣)
    text_rect = fitz.Rect(qr_rect.x1 + qr_margin, key_rect[1], key_rect[2], key_rect[3])
    # -------------------------------------------------

    for index_key in keys:
        try:
            # その都度PDFを開き直す（上書き防止のため）
            doc = fitz.open(input_pdf)
            page = doc[0]

            # フォント登録
            try:
                page.insert_font(fontfile=use_font_path, fontname=font_name)
            except Exception:
                pass

            # QR生成
            qr_json = {"cpk": index_key}
            qr_data = json.dumps(qr_json, ensure_ascii=False)

            qr = qrcode.QRCode(box_size=2, border=0)
            qr.add_data(qr_data)
            qr.make(fit=True)
            qr_img = qr.make_image(fill_color="black", back_color="white")

            img_byte_arr = io.BytesIO()
            qr_img.save(img_byte_arr, format="PNG")
            img_bytes = img_byte_arr.getvalue()

            # 画像挿入
            page.insert_image(qr_rect, stream=img_bytes)

            # テキスト描画
            color_hex = colors.get(index_key.lower(), "#000000")
            icon_char = icons.get(index_key.lower(), "")  # アイコンがない場合は空文字
            rgb_color = hex_to_rgb(color_hex)

            # アイコンがある場合はテキストの前に付与
            display_text = f"{icon_char} {index_key}" if icon_char else index_key

            shape = page.new_shape()
            shape.insert_textbox(
                text_rect,
                display_text,
                fontname=font_name,
                fontsize=10,  # Webclip (pdf_utils.py) は10ptを使用
                color=rgb_color,
                align=0,  # 左寄せ
            )
            shape.commit()

            # --- メタデータ更新 (Synapsen:SkipNormalization の付与) ---
            current_metadata = doc.metadata
            keywords = current_metadata.get("keywords", "")
            skip_flag = "Synapsen:SkipNormalization"

            # フラグが含まれていなければ追加
            if skip_flag not in keywords:
                new_keywords = f"{keywords}; {skip_flag}" if keywords else skip_flag

                # 辞書をコピーして更新し、セットする
                new_metadata = current_metadata.copy()
                new_metadata["keywords"] = new_keywords
                doc.set_metadata(new_metadata)
            # -------------------------------------------------------

            # 保存
            safe_key_name = index_key.replace("/", "_").replace("\\", "_")
            out_filename = f"{input_pdf.stem}_{safe_key_name}.pdf"
            out_path = output_dir / out_filename

            doc.save(str(out_path))
            doc.close()

            print(f"  -> 生成: {out_filename}")

        except Exception as e:
            print(f"  -> エラー ({index_key}): {e}")


def main():
    parser = argparse.ArgumentParser(
        description="PDFにSynapsen用のIndex Key QRコードを埋め込みます。"
    )
    parser.add_argument("input_pdfs", nargs="+", type=Path, help="元となるPDFファイル")

    args = parser.parse_args()

    # 設定読み込み
    script_dir = Path(__file__).parent
    keys, colors, icons, font_path, key_rect = load_synapsen_config(script_dir)

    if not keys:
        print("[エラー] config.ini から Index Key の設定を読み込めませんでした。")
        print(
            "スクリプトを config.ini と同じフォルダ(またはそのサブフォルダ)に配置してください。"
        )
        input("Enterキーを押して終了...")
        sys.exit(1)

    # key_rect が読み込めなかった場合のデフォルト
    if not key_rect:
        key_rect = DEFAULT_KEY_RECT

    for pdf_path in args.input_pdfs:
        if not pdf_path.exists():
            print(f"ファイルが見つかりません: {pdf_path}")
            continue

        # 出力先: 元ファイルと同じ場所に "QR_Output" フォルダを作成
        output_dir = pdf_path.parent / "QR_Output"
        output_dir.mkdir(parents=True, exist_ok=True)

        process_pdf(pdf_path, output_dir, keys, colors, icons, font_path, key_rect)

    print("\nすべての処理が完了しました。")


if __name__ == "__main__":
    main()

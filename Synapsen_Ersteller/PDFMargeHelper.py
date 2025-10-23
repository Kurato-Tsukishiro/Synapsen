
# --- A4サイズ定数 ---
A4_WIDTH = 595.276
A4_HEIGHT = 841.89
IMAGE_DPI = 200


def hex_to_rgb_frac(hex_color):
    """16進数カラーコードをLaTeXのrgb形式(0-1)に変換するヘルパー関数

    Args:
        hex_color (_type_): 16進数のカラーコード

    Returns:
        _type_: RGB形式のカラーコード
    """
    hex_color = hex_color.lstrip('#')
    r, g, b = tuple(int(hex_color[i:i+2], 16) for i in (0, 2, 4))
    return f"{{{r/255:.4f},{g/255:.4f},{b/255:.4f}}}"

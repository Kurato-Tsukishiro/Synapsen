"""
Synapsenアプリケーションテーマ定義。

このモジュールは、Synapsen_Lantcher、Synapsen_Nexus、Synapsen_Ersteller、Synapsen_Normalisierer全体で、
デザインの一貫性を確保するため、色操作用のカラー定数と補助関数を格納します。
"""
from typing import Tuple


# --- Color Constants (Theme Colors) ---
# ==========================================
# Synapsen Identity Colors
# ==========================================
# Main
COLOR_HISUI = "#38b48b"
""" 翡翠色 """
COLOR_TETSU = "#005243"
""" 鉄色 """

# Sub
COLOR_MUSHI = "#20604F"
""" 虫襖色 """
COLOR_SUOU = "#9E3D3F"
""" 蘇芳色 """
COLOR_KIKYO = "#585a9c"
""" 桔梗色 """


# ==========================================
# Semantic Color Mapping
# ==========================================
# モジュールカラー
COLOR_NORMALISIERER = COLOR_KIKYO
""" Normalisierer : 桔梗色 """
COLOR_WATCHDOG = COLOR_SUOU
""" Watchdog : 蘇芳色 (警告/監視) """
COLOR_ERSTELLER = COLOR_HISUI
""" Ersteller : 翡翠色 """
COLOR_NEXUS = COLOR_TETSU
""" Nexus : 鉄色 (基盤) """


# --- Color Utility Functions ---
def hex_to_rgb_frac(hex_color: str) -> Tuple[float, float, float]:
    """
    16進カラー文字列をRGBの分数（0.0～1.0）のタプルに変換する。
    ReportLabなどのライブラリで有用。

    Args:
        hex_color (str): 16進カラー文字列（例:  "#38b48b" or "38b48b").

    Returns:
        Tuple[float, float, float]: 0.0から1.0の間で正規化されたRGB値。(unityで言うとColor型)
    """
    hex_color = hex_color.lstrip("#")
    r, g, b = tuple(int(hex_color[i : i + 2], 16) for i in (0, 2, 4))
    return f"{{{r/255:.4f},{g/255:.4f},{b/255:.4f}}}"


def adjust_brightness(hex_color: str, factor: float = 0.8) -> str:
    """
    16進数カラーコードを受け取り、明度を調整したコードを返すヘルパー関数。

    Args:
        hex_color (str): 16進数カラーコード (Hex color string) (例: "#38b48b").
        factor (float): 調整係数。
                        > 1.0 で明るくなる (例: 1.2 = +20%)。
                        < 1.0 で暗くなる (例: 0.8 = -20%)。
                        初期値の0.8は推奨の影色調整値。(hover_colorなどに使用)

    Returns:
        str: 調整済みの16進数カラーコード。
    """
    hex_color = hex_color.lstrip("#")

    # RGBに分解
    r, g, b = tuple(int(hex_color[i : i + 2], 16) for i in (0, 2, 4))

    # 明度調整 (最大255)
    r, g, b = [max(0, min(255, int(c * factor))) for c in (r, g, b)]
    return f"#{r:02x}{g:02x}{b:02x}"

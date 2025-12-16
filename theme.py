"""
Synapsenアプリケーションテーマ定義。

このモジュールは、Synapsen_Lantcher、Synapsen_Nexus、Synapsen_Ersteller、Synapsen_Normalisierer全体で、
デザインの一貫性を確保するため、色操作用のカラー定数と補助関数を格納します。
"""

from typing import Tuple, Final
import colorsys


class Colors:
    """
    Synapsen Identity Colors (基本パレット)
    """

    # Main
    HISUI: Final[str] = "#38b48b"
    """ 翡翠色 """

    TETSU: Final[str] = "#005243"
    """ 鉄色 """

    # Sub
    MUSHI: Final[str] = "#20604F"
    """ 虫襖色 """

    SUOU: Final[str] = "#9E3D3F"
    """ 蘇芳色 """

    KIKYO: Final[str] = "#585a9c"
    """ 桔梗色 """

    # Accent
    MYOSOTIS: Final[str] = "#89c3eb"
    """ 勿忘草色 """

    HIMAWARI: Final[str] = "#fcc800"
    """ 向日葵色 """

    DOVEGREY: Final[str] = "#9e8b8e"
    """ 鳩羽鼠色 """

    KACIIKAESHI: Final[str] = "#203744"
    """ 褐返色 """

    # Base
    WHITE: Final[str] = "#FFFFFF"
    BLACK: Final[str] = "#000000"


class SemanticColors:
    """
    Semantic Color Mapping (意味的定義)
    Colors の色を参照して役割を定義します。
    """

    # --- Helper Methods (色計算) ---
    @staticmethod
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

    @staticmethod
    def adjust_saturation(hex_color: str, factor: float = 1.0) -> str:
        """
        16進数カラーコードを受け取り、彩度を調整したコードを返すヘルパー関数

        Args:
            hex_color (str): 16進数カラーコード (Hex color string) (例: "#38b48b").
            factor (float): 調整係数。
                            < 1.0 で彩度を下げる(くすませる)。
                            > 1.0 で彩度をあげる(鮮やかになる)。
                            0 で完全にグレーになる。

        Returns:
            str: 調整済みの16進数カラーコード。
        """
        # "#" を除去
        hex_color = hex_color.lstrip("#")

        # エラー回避: 文字列長が足りない場合はそのまま返す
        if len(hex_color) != 6:
            return f"#{hex_color}"

        # 1. 16進数 -> RGB (0~255)
        r, g, b = tuple(int(hex_color[i : i + 2], 16) for i in (0, 2, 4))

        # 2. RGB (0~1.0) -> HSV (色相, 彩度, 明度)
        h, s, v = colorsys.rgb_to_hsv(r / 255.0, g / 255.0, b / 255.0)

        # 3. 彩度(Saturation)を調整 (0.0 ~ 1.0 の範囲にクリップ)
        s = max(0.0, min(1.0, s * factor))

        # 4. HSV -> RGB (0~1.0)
        r, g, b = colorsys.hsv_to_rgb(h, s, v)

        # 5. RGB -> 16進数文字列
        return f"#{int(r*255):02x}{int(g*255):02x}{int(b*255):02x}"

    @staticmethod
    def blend_colors(fg_hex: str, bg_hex: str, alpha: float = 0.5) -> str:
        """
        2つの色を alpha (0.0~1.0) の割合で合成して、擬似的な半透明色を生成する関数。

        Args:
            fg_hex (str): 前景色 (透明にしたい色)
            bg_hex (str): 背景色 (透けて見える後ろの色)
            alpha (float):  不透明度 (1.0で完全な前景色、0.0で完全な背景色)

        Returns:
            str: 調整済みの16進数カラーコード。
        """
        # "#" を除去
        fg_hex = fg_hex.lstrip("#")
        bg_hex = bg_hex.lstrip("#")

        if len(fg_hex) != 6 or len(bg_hex) != 6:
            return f"#{fg_hex}"

        # 1. 16進数 -> RGB
        r1, g1, b1 = tuple(int(fg_hex[i : i + 2], 16) for i in (0, 2, 4))
        r2, g2, b2 = tuple(int(bg_hex[i : i + 2], 16) for i in (0, 2, 4))

        # 2. ブレンド計算: (前景色 * alpha) + (背景色 * (1 - alpha))
        r = int(r1 * alpha + r2 * (1 - alpha))
        g = int(g1 * alpha + g2 * (1 - alpha))
        b = int(b1 * alpha + b2 * (1 - alpha))

        # 3. RGB -> 16進数
        return f"#{r:02x}{g:02x}{b:02x}"

    # --- Helper Methods (変換) ---
    @staticmethod
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

    # --- Module Colors (モジュールごとのメインカラー) ---
    NORMALISIERER: Final[str] = Colors.KIKYO
    """ Normalisierer : 桔梗色 """

    WATCHDOG: Final[str] = Colors.SUOU
    """ Watchdog : 蘇芳色 (警告/監視) """

    ERSTELLER: Final[str] = Colors.HISUI
    """ Ersteller : 翡翠色 """

    NEXUS: Final[str] = Colors.TETSU
    """ Nexus : 鉄色 (基盤) """

    CANVAS: Final[str] = Colors.HIMAWARI
    """ Canvas : 向日葵色 """

    # --- UI Semantics (機能別の色) ---
    UI_BASIC: Final[str] = Colors.MYOSOTIS
    """ 基本的なUIカラー : 勿忘草色 """

    UI_SECONDARY: Final[str] = Colors.KIKYO
    """ 次点のUIカラー : 桔梗色 """

    UI_TERTIARY: Final[str] = Colors.DOVEGREY
    """ 第三のUI要素 : 鳩羽鼠色 """

    # リンク・ジャンプ系
    UI_PREVIEW: Final[str] = Colors.HISUI
    """ 表示系のUIカラー : 翡翠色 """
    UI_LINK: Final[str] = Colors.HISUI
    """ リンク関連機能のUIカラー : 翡翠色 """
    TEXT_LINK: Final[str] = Colors.KIKYO
    """ リンクテキスト(暗) : 桔梗色 """
    TEXT_LINK_BRIGHT: Final[str] = Colors.MYOSOTIS
    """ リンクテキスト(明) : 勿忘草色 """

    # 編集・出力系
    UI_EDIT: Final[str] = Colors.TETSU
    """ 編集関連機能のUI : 鉄色 """
    UI_EXPORT: Final[str] = Colors.TETSU
    """ 出力関連機能のUI : 鉄色 """

    # ラベル・通知系
    LABEL_DENGER: Final[str] = Colors.SUOU
    """ 危険ラベル警告 : 蘇芳色 """
    LABEL_WARNING: Final[str] = adjust_brightness(Colors.HIMAWARI)
    """ 注意ラベル警告 : 向日葵色 (20% 暗) """
    LABEL_INFO: Final[str] = Colors.MYOSOTIS
    """ 情報ラベル警告 : 勿忘草色 """
    LABEL_SUCCESS: Final[str] = Colors.HISUI
    """ 成功ラベル : 翡翠色 """

    # --- Backgrounds ---
    BACKGROUND_PANEL: Final[str] = blend_colors(
        adjust_saturation(Colors.DOVEGREY, 0.7), Colors.WHITE, alpha=0.35
    )
    """ パネル背景色: 薄い灰色 (鳩羽鼠色をベースに彩度を落とし、白とブレンド) """

    BACKGROUND_HOLLOW: Final[str] = blend_colors(
        Colors.WHITE, BACKGROUND_PANEL, 0.4
    )
    """
    窪み背景色 : 非常に薄い灰色 (鳩羽色及びパネル背景色をベースに明るく調整)
    """

    BACKGROUND_DARK_PANEL = Colors.KACIIKAESHI
    """ パネル背景色(DARK) """

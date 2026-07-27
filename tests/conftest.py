# import pytest  # TODO: プッシュ時にはコメントアウト
import sys
import os
from pathlib import Path
import base64

# プロジェクトのルートディレクトリをパスに追加
sys.path.append(os.path.abspath(os.path.join(os.path.dirname(__file__), "..")))


# ========== プロジェクト内モジュールのインポート ===============
from Synapsen_Nexus import utils  # noqa: E402 (インポート例)
from theme import SemanticColors as Colors  # noqa: E402

# ============================================================

# テスト実行コマンド => pytest -s tests\conftest.py


def test_import_verification():
    """
    モジュールが正しくインポートできるかを確認する基本的なテスト。
    """
    # 実際に utils がインポートできているか (Noneでないか) を検証
    assert utils is not None


def test_check_colorcode():
    """
    カラーコードをハードコードする必要がある場合、
    計算したカラーコードをSynapsenを起動せずに確認する。
    """
    color = Colors.UI_BASIC
    print("\n")
    print(f"元の色　: {color}")
    print(f"明度調整: {Colors.adjust_brightness(color)}")
    print(f"彩度調整: {Colors.adjust_saturation(color)}")


def test_get_base64_data():
    """
    画像をbase64に変換する為のコード。
    export_manager.pyに存在する``get_base64_data``と同じだが、これをSynapsenを起動せずに実行する物。
    """
    icon_path_text = r"assets\synapsen.ico"
    icon_path = Path(icon_path_text)

    if not icon_path.exists():
        print("FILE NOT FOUND")
    else:
        with open(icon_path, "rb") as img_file:
            encoded_string = base64.b64encode(img_file.read()).decode("utf-8")

        print("\n")
        print(f"encoded_string = {encoded_string}")

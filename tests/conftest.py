# import pytest  # TODO: プッシュ時にはコメントアウト
import sys
import os

# プロジェクトのルートディレクトリをパスに追加
sys.path.append(os.path.abspath(os.path.join(os.path.dirname(__file__), "..")))


# ========== プロジェクト内モジュールのインポート ===============
from Synapsen_Nexus import utils  # noqa: E402 (インポート例)
from theme import SemanticColors as Colors  # noqa: E402

# ============================================================


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
    # テスト実行コマンド => pytest -s tests\conftest.py
    color = Colors.UI_BASIC
    print("\n")
    print(f"元の色　: {color}")
    print(f"明度調整: {Colors.adjust_brightness(color)}")
    print(f"彩度調整: {Colors.adjust_saturation(color)}")

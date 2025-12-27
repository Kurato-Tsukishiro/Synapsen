import pytest  # noqa: F401
import sys
import os

# プロジェクトのルートディレクトリをパスに追加
sys.path.append(os.path.abspath(os.path.join(os.path.dirname(__file__), "..")))

# ========== プロジェクト内モジュールのインポート ===============
from Synapsen_Nexus import utils  # noqa: E402 (インポート例)

# ============================================================


def test_import_verification():
    """
    モジュールが正しくインポートできるかを確認する基本的なテスト。
    """
    # 実際に utils がインポートできているか (Noneでないか) を検証
    assert utils is not None

import customtkinter as ctk
import sys
import subprocess
import traceback
from pathlib import Path
from PIL import Image
from tkinter import messagebox

# --- パス設定 (exe/script両対応) ---

# 1. 実行環境のルートディレクトリを特定
if getattr(sys, "frozen", False):
    # 【EXE実行時】
    # 物理的なexeの場所 (config.ini, assetsの読み込み用)
    BASE_DIR = Path(sys.executable).parent
    # 内部ファイルが展開されている一時フォルダ (モジュールのインポート用)
    INTERNAL_DIR = Path(sys._MEIPASS)
else:
    # 【スクリプト実行時】
    BASE_DIR = Path(__file__).parent
    INTERNAL_DIR = BASE_DIR

# パスを通す
if str(INTERNAL_DIR) not in sys.path:
    sys.path.insert(0, str(INTERNAL_DIR))

# 3. 各ツールのサブフォルダも sys.path に追加
# これにより、各ツール内の "import dnd_window" 等のローカルインポートが解決されます
tool_subdirs = ["Synapsen_Normalisierer", "Synapsen_Ersteller", "Synapsen_Nexus"]

for subdir in tool_subdirs:
    # EXE内では INTERNAL_DIR 配下にフォルダが存在します
    target_path = INTERNAL_DIR / subdir
    # パスに追加 (存在チェックは緩める: PyInstallerのバンドル構造によってはフォルダとして見えない場合もあるため)
    sys.path.insert(0, str(target_path))

# ログ設定 (共通)
try:
    from logging_setup import setup_logging

    setup_logging("Synapsen_Launcher")
    import logging

    logger = logging.getLogger("Launcher")
except ImportError:
    import logging

    logger = logging.getLogger()

# --- PyInstaller スプラッシュスクリーン制御 ---
# ビルド環境以外(スクリプト実行時)ではインポートエラーになるためtryで囲む
try:
    import pyi_splash
except ImportError:
    pyi_splash = None

# --- Synapsen Theme Colors ---
COLOR_HISUI = "#38b48b"  # 翡翠色 (Main: Ersteller)
COLOR_TETSU = "#005243"  # 鉄色   (Main: Nexus)
COLOR_MUSHI = "#20604F"  # 虫襖   (Sub)
COLOR_SUOU = "#9E3D3F"   # 蘇芳色 (Sub: Watchdog)
COLOR_KIKYO = "#585a9c"  # 桔梗色 (Sub: Normalisierer)


class SynapsenLauncher(ctk.CTk):
    def __init__(self):
        super().__init__()

        self.title("Synapsen Launcher")
        self.geometry("600x550")

        # 物理パスを使用 (assets等はexeの隣にある前提)
        self.base_path = BASE_DIR

        # アイコン設定
        self.icon_path = (self.base_path / "assets" / "synapsen.ico").resolve()
        if self.icon_path.exists():
            try:
                self.iconbitmap(default=str(self.icon_path))
            except Exception:
                pass

        self.grid_columnconfigure(0, weight=1)
        self.grid_rowconfigure(0, weight=1)

        # UI構築
        self._create_ui()

    def _create_ui(self):
        main_frame = ctk.CTkFrame(self)
        main_frame.grid(row=0, column=0, padx=20, pady=20, sticky="nsew")
        main_frame.grid_columnconfigure((0, 1), weight=1)

        # バナー画像 (メインUI用)
        banner_path = (self.base_path / "assets" / "synapsen_banner.png").resolve()
        title_widget = None
        if banner_path.exists():
            try:
                pil_image = Image.open(banner_path)
                target_width = 400
                w_percent = target_width / float(pil_image.size[0])
                target_height = int((float(pil_image.size[1]) * float(w_percent)))
                banner_image = ctk.CTkImage(
                    light_image=pil_image,
                    dark_image=pil_image,
                    size=(target_width, target_height),
                )
                title_widget = ctk.CTkLabel(main_frame, text="", image=banner_image)
            except Exception:
                pass

        # 画像がない、または読み込み失敗時はテキストで表示
        if title_widget is None:
            title_widget = ctk.CTkLabel(
                main_frame, text="Synapsen Toolset", font=("Arial", 24, "bold")
            )

        title_widget.grid(row=0, column=0, columnspan=2, pady=(30, 20))

        # ツール定義
        tools = [
            (
                "Normalisierer",
                "正規化: PDFの整形・OCR・Webクリップ",
                "--normalisierer",
                COLOR_KIKYO,  # 桔梗色
                False,
            ),
            (
                "Ersteller",
                "統合: メタデータ編集・月次PDF作成",
                "--ersteller",
                COLOR_HISUI,  # 翡翠色
                False,
            ),
            (
                "Nexus",
                "閲覧: 検索・ネットワーク思考",
                "--nexus",
                COLOR_TETSU,  # 鉄色 (基盤)
                False,
            ),
            (
                "Watchdog",
                "監視: 自動正規化(常駐)",
                "--watchdog",
                COLOR_SUOU,  # 蘇芳色 (警告/監視)
                False,
            ),
        ]

        for i, (name, desc, arg_flag, color, use_console) in enumerate(tools):
            row = i // 2 + 1
            col = i % 2

            btn_frame = ctk.CTkFrame(main_frame, fg_color="transparent")
            btn_frame.grid(row=row, column=col, padx=10, pady=10, sticky="nsew")

            btn = ctk.CTkButton(
                btn_frame,
                text=name,
                font=("Arial", 16, "bold"),
                height=50,
                fg_color=color,
                hover_color=self._adjust_brightness(color, 0.8),
                command=lambda f=arg_flag, c=use_console: self.launch_self(f, c),
            )
            btn.pack(fill="x", pady=(0, 5))

            ctk.CTkLabel(
                btn_frame, text=desc, font=("Arial", 11), text_color="gray"
            ).pack()

        exit_btn = ctk.CTkButton(
            self, text="終了", fg_color="gray", width=100, command=self.destroy
        )
        exit_btn.grid(row=(len(tools) + 1) // 2 + 1, column=0, columnspan=2, pady=20)

    def _adjust_brightness(self, hex_color, factor=0.8):
        """
        16進数カラーコードを受け取り、明度を調整したコードを返すヘルパー関数
        factor < 1.0 で暗く、> 1.0 で明るくなります。
        """
        hex_color = hex_color.lstrip("#")

        # RGBに分解
        r, g, b = tuple(int(hex_color[i : i + 2], 16) for i in (0, 2, 4))

        # 明度調整 (最大255)
        r, g, b = [max(0, min(255, int(c * factor))) for c in (r, g, b)]
        return f"#{r:02x}{g:02x}{b:02x}"

    def launch_self(self, flag, use_console=False):
        """自分自身を別プロセスとして起動"""
        try:
            # 実行パスの解決
            if getattr(sys, "frozen", False):
                # EXE実行時: Synapsen.exe --flag
                executable = sys.executable
                cmd = [executable, flag]
            else:
                # スクリプト実行時: python Synapsen_Launcher.py --flag
                executable = sys.executable
                script_path = str(Path(__file__).resolve())
                cmd = [executable, script_path, flag]

            # cwdは物理的なベースパスを指定
            cwd_path = str(self.base_path.resolve())

            if sys.platform == "win32":
                creation_flags = (
                    subprocess.CREATE_NEW_CONSOLE if use_console else 0x08000000
                )
                subprocess.Popen(cmd, cwd=cwd_path, creationflags=creation_flags)
            else:
                subprocess.Popen(cmd, cwd=cwd_path)

            logger.info(f"Launched sub-process: {cmd}")
        except Exception as e:
            logger.error(f"Launch failed: {e}")
            messagebox.showerror("起動エラー", f"ツールの起動に失敗しました:\n{e}")


# --- スプラッシュを閉じるヘルパー関数 ---
def close_splash():
    if pyi_splash and pyi_splash.is_alive():
        try:
            pyi_splash.close()
        except Exception:
            pass


# --- メインエントリーポイント ---
if __name__ == "__main__":
    ctk.set_appearance_mode("System")

    # 引数チェックによる分岐
    # ※ 各ツールのインポートはここで初めて行う (遅延インポートによる高速化)
    if len(sys.argv) > 1:
        mode = sys.argv[1]

        # モジュール変数が定義されているか確認してから起動
        try:
            if mode == "--normalisierer":
                # 各ツール起動時も、準備ができたらスプラッシュを閉じる
                # (サブプロセス起動時にもスプラッシュが出る場合があるため念のため)
                close_splash()
                from Synapsen_Normalisierer.Synapsen_Normalisierer_main import (
                    Synapsen_Normalisierer,
                )

                app = Synapsen_Normalisierer()
                if (
                    hasattr(app, "iconbitmap")
                    and (BASE_DIR / "assets" / "synapsen.ico").exists()
                ):
                    try:
                        app.iconbitmap(
                            default=str(BASE_DIR / "assets" / "synapsen.ico")
                        )
                    except Exception:
                        pass
                app.mainloop()

            elif mode == "--ersteller":
                close_splash()
                from Synapsen_Ersteller.Synapsen_Ersteller_main import (
                    Synapsen_Ersteller,
                )

                app = Synapsen_Ersteller()
                if (
                    hasattr(app, "iconbitmap")
                    and (BASE_DIR / "assets" / "synapsen.ico").exists()
                ):
                    try:
                        app.iconbitmap(
                            default=str(BASE_DIR / "assets" / "synapsen.ico")
                        )
                    except Exception:
                        pass
                app.mainloop()

            elif mode == "--nexus":
                close_splash()
                from Synapsen_Nexus.Synapsen_Nexus_main import Synapsen_Nexus

                app = Synapsen_Nexus()
                if (
                    hasattr(app, "iconbitmap")
                    and (BASE_DIR / "assets" / "synapsen.ico").exists()
                ):
                    try:
                        app.iconbitmap(
                            default=str(BASE_DIR / "assets" / "synapsen.ico")
                        )
                    except Exception:
                        pass
                app.mainloop()

            elif mode == "--watchdog":
                close_splash()
                from Synapsen_Normalisierer import Synapsen_Watchdog

                Synapsen_Watchdog.main()

            else:
                # 不明な引数の場合はランチャー
                app = SynapsenLauncher()
                # ランチャーの準備が整ったらスプラッシュを閉じる
                close_splash()
                app.mainloop()

        except Exception as e:
            close_splash()  # エラー時も閉じる
            messagebox.showerror(
                "実行エラー",
                f"アプリの実行中にエラーが発生しました:\n{e}\n\n{traceback.format_exc()}",
            )
    else:
        # 引数なし -> ランチャー起動
        app = SynapsenLauncher()
        close_splash()
        app.mainloop()

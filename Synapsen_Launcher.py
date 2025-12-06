import customtkinter as ctk
import sys
import subprocess
from pathlib import Path
from PIL import Image  # 【追加】画像処理用

# ログ設定の読み込み
try:
    from logging_setup import setup_logging

    setup_logging("Synapsen_Launcher")
    import logging

    logger = logging.getLogger("Launcher")
except ImportError:
    import logging

    logger = logging.getLogger()

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
        self.geometry("600x500")  # 画像が入る分、少し高さを広げました

        if getattr(sys, "frozen", False):
            self.base_path = Path(sys.executable).parent
        else:
            self.base_path = Path(__file__).parent

        self.icon_path = self.base_path / "assets" / "synapsen.ico"
        if self.icon_path.exists():
            try:
                self.iconbitmap(default=str(self.icon_path))
            except Exception:
                pass

        self.grid_columnconfigure(0, weight=1)
        self.grid_rowconfigure(0, weight=1)

        self._create_ui()

    def _create_ui(self):
        main_frame = ctk.CTkFrame(self)
        main_frame.grid(row=0, column=0, padx=20, pady=20, sticky="nsew")
        main_frame.grid_columnconfigure((0, 1), weight=1)

        # --- タイトルバナー画像の表示 ---
        # 1. 画像パスの設定 (PNG形式を使用)
        banner_path = self.base_path / "assets" / "synapsen_banner.png"

        title_widget = None

        if banner_path.exists():
            try:
                # 2. 画像の読み込み
                pil_image = Image.open(banner_path)

                # 3. 表示サイズの設定 (アスペクト比を維持しつつ幅400pxに合わせる例)
                target_width = 400
                w_percent = target_width / float(pil_image.size[0])
                target_height = int((float(pil_image.size[1]) * float(w_percent)))

                banner_image = ctk.CTkImage(
                    light_image=pil_image,
                    dark_image=pil_image,
                    size=(target_width, target_height),
                )

                # 4. 画像ラベルの作成 (textは空にする)
                title_widget = ctk.CTkLabel(main_frame, text="", image=banner_image)
            except Exception as e:
                logger.error(f"Failed to load banner image: {e}")

        # 画像がない、または読み込み失敗時はテキストで表示
        if title_widget is None:
            title_widget = ctk.CTkLabel(
                main_frame, text="Synapsen Toolset", font=("Arial", 24, "bold")
            )

        title_widget.grid(row=0, column=0, columnspan=2, pady=(30, 20))

        # --- ボタンの定義 ---
        # (ボタン名, 説明, スクリプトパス, 色, コンソール表示有無)
        tools = [
            (
                "Normalisierer",
                "正規化: PDFの整形・OCR・Webクリップ",
                "Synapsen_Normalisierer/Synapsen_Normalisierer_main.py",
                COLOR_KIKYO,  # 桔梗色
                False,
            ),
            (
                "Ersteller",
                "統合: メタデータ編集・月次PDF作成",
                "Synapsen_Ersteller/Synapsen_Ersteller_main.py",
                COLOR_HISUI,  # 翡翠色
                False,
            ),
            (
                "Nexus",
                "閲覧: データベース検索・ネットワーク思考",
                "Synapsen_Nexus/Synapsen_Nexus_main.py",
                COLOR_TETSU,  # 鉄色 (基盤)
                False,
            ),
            (
                "Watchdog",
                "監視: フォルダ監視による自動正規化",
                "Synapsen_Normalisierer/Synapsen_Watchdog.py",
                COLOR_SUOU,  # 蘇芳色 (警告/監視)
                True,
            ),
        ]

        for i, (name, desc, script_rel_path, color, use_console) in enumerate(tools):
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
                command=lambda p=script_rel_path, c=use_console: self.launch_tool(p, c),
            )
            btn.pack(fill="x", pady=(0, 5))

            desc_label = ctk.CTkLabel(
                btn_frame, text=desc, font=("Arial", 11), text_color="gray"
            )
            desc_label.pack()

        exit_btn = ctk.CTkButton(
            self, text="終了", fg_color="gray", width=100, command=self.destroy
        )
        # 行番号を動的に調整 (ツール行数 + 1)
        exit_row = (len(tools) + 1) // 2 + 1
        exit_btn.grid(row=exit_row, column=0, columnspan=2, pady=20)

    def _adjust_brightness(self, hex_color, factor=0.8):
        """
        16進数カラーコードを受け取り、明度を調整したコードを返すヘルパー関数
        factor < 1.0 で暗く、> 1.0 で明るくなります。
        """
        hex_color = hex_color.lstrip("#")

        # RGBに分解
        r = int(hex_color[0:2], 16)
        g = int(hex_color[2:4], 16)
        b = int(hex_color[4:6], 16)

        # 明度調整 (最大255)
        r = max(0, min(255, int(r * factor)))
        g = max(0, min(255, int(g * factor)))
        b = max(0, min(255, int(b * factor)))

        return f"#{r:02x}{g:02x}{b:02x}"

    def launch_tool(self, script_relative_path, use_console=False):
        """
        指定されたスクリプトを別プロセスで起動する。
        """
        script_path = self.base_path / script_relative_path

        if not script_path.exists():
            print(f"Error: Script not found at {script_path}")
            return

        try:
            if sys.platform == "win32":
                creation_flags = 0
                if use_console:
                    creation_flags = subprocess.CREATE_NEW_CONSOLE
                else:
                    # CREATE_NO_WINDOW
                    creation_flags = 0x08000000

                subprocess.Popen(
                    [sys.executable, str(script_path)],
                    cwd=str(self.base_path),
                    creationflags=creation_flags,
                )
            else:
                subprocess.Popen(
                    [sys.executable, str(script_path)], cwd=str(self.base_path)
                )

            logger.info(f"Launched: {script_relative_path} (Console: {use_console})")

        except Exception as e:
            logger.error(f"Failed to launch {script_relative_path}: {e}")


if __name__ == "__main__":
    ctk.set_appearance_mode("System")
    app = SynapsenLauncher()
    app.mainloop()

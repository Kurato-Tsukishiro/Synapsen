# === 1. 標準ライブラリ ===
import os
import sys
from pathlib import Path
import tkinter
import csv
import json
from textwrap import dedent
import shutil
import tempfile
import configparser
import sqlite3
import logging

# === 2. プロジェクトルートをパスに追加 ===
current_dir = Path(__file__).parent
root_dir = current_dir.parent
if str(root_dir) not in sys.path:
    sys.path.append(str(root_dir))

# === 3. プロジェクト内モジュールとサードパーティ ===
import db_recovery_tool                         # noqa: E402
from tkinter import messagebox                  # noqa: E402 (tkinterグループ)
import customtkinter as ctk                     # noqa: E402
import pandas as pd                             # noqa: E402
import PDFMargeHelper as Helper                 # noqa: E402
import pdf_processor as Process                 # noqa: E402
import reportlab_generator as Generator         # noqa: E402
import gui_dialogs as Dialogs                   # noqa: E402
from pypdf import PdfReader, PdfWriter          # noqa: E402
from Synapsen_Nexus import utils as NexusUtils  # noqa: E402
from logging_setup import setup_logging  # noqa: E402

# ==============================================================================
# 定数定義
# ==============================================================================
# 現在のDBのスキーマバージョン
CURRET_SCHEMA_VERSION = 1.1

# ==============================================================================
# ロギング設定の初期化
# ==============================================================================
# 親ディレクトリ(ルート)をパスに追加して logging_setup.py をインポート可能にする
current_dir = Path(__file__).parent
root_dir = current_dir.parent
if str(root_dir) not in sys.path:
    sys.path.append(str(root_dir))

try:
    # アプリ名を指定して初期化
    setup_logging("Synapsen_Ersteller")
    logger = logging.getLogger("Ersteller")  # このファイル用のロガー取得
except ImportError:
    # logging_setup.py がない場合のフォールバック（print出力）
    print("Warning: logging_setup.py not found. Logging disabled.")

    class MockLogger:
        def info(self, msg):
            print(f"[INFO] {msg}")

        def error(self, msg, exc_info=None):
            print(f"[ERROR] {msg} {exc_info if exc_info else ''}")

        def warning(self, msg):
            print(f"[WARN] {msg}")

    logger = MockLogger()
# ==============================================================================


# ==============================================================================
# メインアプリケーションクラス
# ==============================================================================
class Synapsen_Ersteller(ctk.CTk):
    def __init__(self):
        super().__init__()
        self.icon_path = self.get_icon_path()
        self.title("Synapse Ersteller")
        self.geometry("800x700")
        self.grid_columnconfigure(0, weight=1)
        # ボタン用に row=1, 2 を確保し、リスト用に row=3 の weight を 1 に設定
        self.grid_rowconfigure(3, weight=1)

        # 用紙サイズを保持する変数
        self.paper_size = "A4"  # デフォルト
        self.paper_width = Helper.A4_WIDTH
        self.paper_height = Helper.A4_HEIGHT

        self.load_config()

        self.all_notes_info = []
        self.predefined_tags = []
        self.load_predefined_tags()

        # 一括編集機能のための選択状態保持
        self.selected_notes = set()  # 選択されたノートの 'key' を保持する

        self.label = ctk.CTkLabel(
            self,
            text="Synapsen Normalisiererで処理済みのフォルダを読み込んでください。",
        )
        self.label.grid(row=0, column=0, padx=10, pady=10, sticky="ew")

        # --- 1段目のボタフレーム (ファイル操作 + PDF生成 + 設定) ---
        top_button_frame = ctk.CTkFrame(self, fg_color="transparent")
        top_button_frame.grid(row=1, column=0, padx=10, pady=5, sticky="ew")

        # フォルダ/CSV操作 (左側)
        load_frame = ctk.CTkFrame(top_button_frame, fg_color="transparent")
        load_frame.pack(side="left", padx=0)
        ctk.CTkButton(
            load_frame, text="フォルダから新規読み込み", command=self.scan_folder
        ).pack(side="left", padx=5)
        ctk.CTkButton(
            load_frame, text="リスト読込 (CSV)", command=self.load_from_csv
        ).pack(side="left", padx=5)
        ctk.CTkButton(
            load_frame, text="リストをフォルダと同期", command=self.sync_with_folder
        ).pack(side="left", padx=5)
        ctk.CTkButton(
            load_frame, text="リスト保存 (CSV)", command=self.save_to_csv
        ).pack(side="left", padx=5)

        # === 右側のボタン群 (PDF生成 / DB復旧 / 設定) ===
        # pack(side="right") なので、コード上で先に書いたものが「より右側」に配置されます

        # [1] PDF生成 (一番右)
        ctk.CTkButton(
            top_button_frame,
            text="統合PDFを生成",
            command=self.generate_pdf,
            fg_color="green",
            hover_color="darkgreen",
        ).pack(side="right", padx=10)

        # [2] DB復旧ツール (その左)
        ctk.CTkButton(
            top_button_frame,
            text="DB復旧ツール",
            command=self.open_recovery_tool,
            fg_color="#17a2b8",
            hover_color="#138496",  # シアン系
            width=100,
        ).pack(side="right", padx=5)

        # [3] 設定ボタン
        ctk.CTkButton(
            top_button_frame,
            text="設定 (Config)",
            command=self.open_config_editor,
            fg_color="#555555",  # ツール系なのでグレー
            hover_color="#333333",
            width=80,
        ).pack(side="right", padx=5)

        # --- 2段目のボタフレーム (一括編集) ---
        batch_button_frame = ctk.CTkFrame(self, fg_color="transparent")
        batch_button_frame.grid(row=2, column=0, padx=10, pady=0, sticky="ew")

        self.batch_edit_button = ctk.CTkButton(
            batch_button_frame,
            text="一括編集 (0)",
            command=self.open_batch_editor,
            state="disabled",
            fg_color="#585a9c",  # 桔梗色
            hover_color="#494B83",  # 濃い桔梗色
        )
        self.batch_edit_button.pack(side="left", padx=5)

        self.deselect_all_button = ctk.CTkButton(
            batch_button_frame,
            text="選択解除",
            command=self.deselect_all,
            state="disabled",
            fg_color="#6C757D",  # セカンダリ・グレー
            hover_color="#5A6268",  # 濃いグレー
        )
        self.deselect_all_button.pack(side="left", padx=5)

        self.copy_links_button = ctk.CTkButton(
            batch_button_frame,
            text="リンクコピー (0)",
            command=self.copy_selected_links,
            state="disabled",
            fg_color="#28a745",
            hover_color="#218838",
        )
        self.copy_links_button.pack(side="left", padx=5)

        self.scrollable_frame = ctk.CTkScrollableFrame(self, label_text="読み込み結果")
        self.scrollable_frame.grid(row=3, column=0, padx=10, pady=10, sticky="nsew")

    def get_icon_path(self):
        """
        実行環境(.exe or .py)に応じて、
        プロジェクトルートの 'assets' フォルダにある
        'synapsen.ico' のパスを返す。
        """
        try:
            if getattr(sys, "frozen", False):
                # .exe実行の場合 (exeと同じフォルダがプロジェクトルート)
                project_root = Path(sys.executable).parent
            else:
                # .pyスクリプト実行の場合 (このファイルの親フォルダがプロジェクトルート)
                project_root = Path(__file__).parent.parent

            icon_path = project_root / "assets" / "synapsen.ico"

            if icon_path.is_file():
                return icon_path
        except Exception as e:
            logger.error(f"Error finding icon path: {e}")
        return None

    def load_config(self):
        # 1. 実行ファイルの場所を基準としたbase_pathを最初に定義します
        if getattr(sys, "frozen", False):
            # .exe実行の場合
            base_path = os.path.dirname(sys.executable)
        else:
            # .pyスクリプト実行の場合
            base_path = os.path.dirname(os.path.abspath(__file__))

        # 2. .exeか.pyかでconfig.iniの場所を決定します
        if getattr(sys, "frozen", False):
            # .exe実行の場合（config.ini は .exe と同じフォルダ）
            config_path = os.path.join(base_path, "config.ini")
        else:
            # スクリプト実行の場合（config.ini は .py の1つ上のフォルダ）
            config_path = os.path.join(
                os.path.abspath(os.path.join(base_path, "..")), "config.ini"
            )
        logger.debug(f"Loading config from: {config_path}")

        # config.ini があるフォルダのパスを基準として定義
        config_dir = os.path.dirname(config_path)

        config = configparser.ConfigParser(interpolation=None)

        # 3. configファイルが存在しない場合の処理
        if not os.path.exists(config_path):
            config["Paths"] = {
                "tags_data_path": "tags.txt",
                "font_path": r"C:\windows\fonts\msgothic.ttc",
            }
            config["ReportLab"] = {
                "paper_size": "A4",
                "font": "MS UI Gothic",
                "author": "Your Name",
                "title_prefix": "月刊 統合ノート",
            }
            config["CommonplaceKeys"] = {
                "options": "タスク,アイデア,思考・考察,コミュニケーション,学習・情報収集,日常・その他"
            }
            config["Extraction"] = {"key_rect": "26, 13, 400, 73"}
            config["KeyIcons"] = {
                "タスク": "♥",
                "アイデア": "♥",
                "思考・考察": "♥",
                "コミュニケーション": "♥",
                "学習・情報収集": "♥",
                "日常・その他": "♥",
            }
            config["KeyColors"] = {
                "タスク": "FE0000",
                "アイデア": "FFFF02",
                "思考・考察": "8802FF",
                "コミュニケーション": "02FF01",
                "学習・情報収集": "02FFFF",
                "日常・その他": "F2F2F2",
            }
            with open(config_path, "w", encoding="utf-8") as f:
                config.write(f)

        config.read(config_path, encoding="utf-8")

        # 4. [Paths] font_path の読み込み
        font_path_from_config = config.get("Paths", "font_path", fallback="")
        expanded_path = os.path.expandvars(font_path_from_config)

        if os.path.isabs(expanded_path):
            # configの値が絶対パス（または環境変数展開後、絶対パスになった）の場合、そのまま使用
            self.font_path = expanded_path
            logger.debug(f"Font path is ABSOLUTE: {self.font_path}")
        else:
            # configの値が相対パスの場合、config_dir と結合する
            self.font_path = os.path.join(config_dir, expanded_path)
            logger.debug(f"Font path is RELATIVE, resolved to: {self.font_path}")

        # 5. tags_data_pathの解決
        tags_path_from_config = config.get(
            "Paths", "tags_data_path", fallback="tags.txt"
        )
        expanded_path = os.path.expandvars(tags_path_from_config)
        if os.path.isabs(expanded_path):
            self.tags_data_path = expanded_path
        else:
            # configの値が相対パスの場合、config_dir と結合する
            self.tags_data_path = os.path.join(config_dir, tags_path_from_config)

        # 6. default_db_path の解決
        default_db_path_str = config.get("Paths", "database_path", fallback="")
        expanded_path = os.path.expandvars(default_db_path_str)
        if not expanded_path:
            self.default_db_path = None
            logger.debug("config.ini [Paths][database_path] が未設定です。")
        elif os.path.isabs(expanded_path):
            self.default_db_path = expanded_path
        else:
            # configの値が相対パスの場合、config_dir と結合する
            self.default_db_path = os.path.join(config_dir, expanded_path)

        # 7. Automation設定の読み込み
        # 自動結合設定の読み込み
        self.auto_append_db = config.getboolean(
            "Automation", "auto_append_to_default_db", fallback=False
        )
        if self.auto_append_db and not self.default_db_path:
            logger.warning(
                "警告: auto_append_to_default_db が True ですが、"
                "database_path が未設定のため無効化されます。"
            )
            self.auto_append_db = False

        # 個別出力設定の読み込み
        self.create_individual_csv = config.getboolean(
            "Automation", "create_individual_csv", fallback=False
        )

        section_name = "ReportLab"

        self.paper_size = config.get(section_name, "paper_size", fallback="A4").upper()
        if self.paper_size == "A5":
            self.paper_width = Helper.A5_WIDTH
            self.paper_height = Helper.A5_HEIGHT
            logger.debug("Ersteller paper size set to A5")
        else:
            self.paper_size = "A4"
            self.paper_width = Helper.A4_WIDTH
            self.paper_height = Helper.A4_HEIGHT
            logger.debug("Ersteller paper size set to A4")

        # フォントパスの環境変数展開
        font_from_config = config.get(section_name, "font", fallback="").strip()
        raw_font = os.path.expandvars(font_from_config)

        if raw_font:
            expanded_reportlab_font = os.path.expandvars(raw_font)
        else:
            expanded_reportlab_font = ""

        if not expanded_reportlab_font:
            # 空文字の場合そのまま使用
            self.reportlab_font = ""
        elif os.path.isabs(expanded_reportlab_font):
            # configの値が絶対パス（または環境変数展開後、絶対パスになった）の場合、そのまま使用
            self.reportlab_font = expanded_reportlab_font
        else:
            # configの値が相対パスの場合、config_dir と結合する
            self.reportlab_font = os.path.join(config_dir, expanded_path)

        self.reportlab_author = config.get(section_name, "author", fallback="Your Name")
        self.reportlab_title_prefix = config.get(
            section_name, "title_prefix", fallback="月刊 統合ノート"
        )
        try:
            self.max_pages_per_volume = config.getint(
                section_name, "max_pages", fallback=0
            )
        except ValueError:
            self.max_pages_per_volume = 0

        # その他 (CommonplaceKeysなど)
        self.commonplace_key_options = [
            opt.strip()
            for opt in config.get("CommonplaceKeys", "options", fallback="").split(",")
        ]
        rect_str = config.get("Extraction", "key_rect", fallback="0,0,0,0").split(",")
        self.key_rect = tuple(map(float, rect_str))
        self.key_icons = (
            {k.lower(): v for k, v in config.items("KeyIcons")}
            if config.has_section("KeyIcons")
            else {}
        )
        self.key_colors = (
            {k.lower(): v for k, v in config.items("KeyColors")}
            if config.has_section("KeyColors")
            else {}
        )

    def load_predefined_tags(self):
        try:
            tag_file = Path(self.tags_data_path)
            if tag_file.is_file():
                with open(tag_file, "r", encoding="utf-8") as f:
                    self.predefined_tags = [
                        line.strip()
                        for line in f
                        if line.strip() and not line.startswith("#")
                    ]
                logger.info(
                    f"{len(self.predefined_tags)}件の事前定義タグを読み込みました。"
                )
                logger.debug(f"{self.tags_data_path}")
        except Exception as e:
            logger.error(f"tags.txtの読み込み中にエラーが発生しました: {e}")

    def save_to_csv(self):
        if not self.all_notes_info:
            return
        filepath = tkinter.filedialog.asksaveasfilename(
            defaultextension=".csv",
            filetypes=[("CSV", "*.csv")],
            title="CSVファイルを保存",
        )
        if not filepath:
            return
        try:
            with open(filepath, "w", newline="", encoding="utf-8-sig") as f:
                header = [
                    "date",
                    "time",
                    "title",
                    "pages",
                    "tags",
                    "key",
                    "memo",
                    "commonplace_key",
                    "filepath",
                    "summary",
                ]
                writer = csv.DictWriter(f, fieldnames=header, extrasaction="ignore")
                writer.writeheader()
                for note in self.all_notes_info:
                    note_to_write = note.copy()
                    note_to_write["tags"] = ";".join(
                        sorted(note_to_write.get("tags", []))
                    )
                    writer.writerow(note_to_write)
            self.label.configure(text=f"保存完了: {os.path.basename(filepath)}")
        except Exception as e:
            self.label.configure(text=f"エラー: 保存失敗 - {e}")

    def load_from_csv(self):
        filepath = tkinter.filedialog.askopenfilename(
            filetypes=[("CSV", "*.csv")], title="CSVファイルを開く"
        )
        if not filepath:
            return
        try:
            new_notes_info = []
            with open(filepath, "r", encoding="utf-8-sig") as f:
                reader = csv.DictReader(f)
                for row in reader:
                    if row.get("tags"):
                        row["tags"] = row.get("tags", "").split(";")
                    else:
                        row["tags"] = []
                    row["pages"] = int(row.get("pages", 0))
                    row["is_warning"] = row.get("date") in ["日付不明", "読み込み失敗"]
                    row["full_text"] = row.get("full_text", "")
                    row["summary"] = row.get("summary", "")
                    new_notes_info.append(row)
            self.all_notes_info = new_notes_info
            self.deselect_all()  # 読み込み時は選択をリセット
            self.update_note_list()
            self.label.configure(text=f"読み込み完了: {os.path.basename(filepath)}")
        except Exception as e:
            self.label.configure(text=f"エラー: 読み込み失敗 - {e}")

    def append_to_master_db(self, notes_to_append):
        """
        マスターDB（config.iniのdatabase_path）に、
        重複をチェックしながらノート情報を追記する。
        """
        if not self.default_db_path:
            logger.error("DBパスが設定されていないため追記できません。")
            return

        master_db_path = Path(self.default_db_path)
        table_name = "notes"

        if not notes_to_append:
            return

        df_new_notes = pd.DataFrame(notes_to_append)
        df_new_notes = df_new_notes[
            df_new_notes["key"].notna() & (df_new_notes["key"] != "")
        ]
        if df_new_notes.empty:
            logger.info("DB追記対象のノート（有効なKeyを持つもの）がありません。")
            return

        # 入力データ内の重複を排除 (Keyが同じなら最初のものを優先)
        df_new_notes = df_new_notes.drop_duplicates(subset=["key"], keep="first")

        # タグリストを ';' 区切りの文字列に変換
        if "tags" in df_new_notes.columns:

            # タグをソートし、';'区切りの文字列に変換する関数を定義
            def format_tags_for_db(tags):
                if isinstance(tags, list):
                    return ";".join(sorted(tags))
                return ""  # リストでない場合は空文字を返す

            # apply に定義した関数を渡す
            df_new_notes["tags"] = df_new_notes["tags"].apply(format_tags_for_db)

        conn = sqlite3.connect(self.default_db_path)
        cursor = conn.cursor()
        try:
            # --- FTS5テーブルとトリガーの作成 ---
            #    (db_recovery_tool.py と全く同じSQL)
            create_table_sql = """
            CREATE TABLE IF NOT EXISTS notes (
                "date" TEXT, "time" TEXT, "title" TEXT, "pages" INTEGER,
                "tags" TEXT, "key" TEXT PRIMARY KEY, "memo" TEXT,
                "commonplace_key" TEXT, "filepath" TEXT, "full_text" TEXT,
                "merged_pdf_filename" TEXT, "merged_start_page" TEXT,
                "summary" TEXT
            )
            """
            cursor.execute(create_table_sql)

            create_fts_sql = """
            CREATE VIRTUAL TABLE IF NOT EXISTS notes_fts USING fts5(
                key, title, memo, tags, full_text,
                summary,
                content='notes', content_rowid='key'
            );
            """
            cursor.execute(create_fts_sql)

            trigger_sql = dedent(
                """
            CREATE TRIGGER IF NOT EXISTS trg_notes_after_insert
                AFTER INSERT ON notes
            BEGIN
                INSERT INTO notes_fts(rowid, key, title,
                                      memo, tags, full_text, summary)
                VALUES (new.key, new.key, new.title,
                        new.memo, new.tags, new.full_text, new.summary);
            END;

            CREATE TRIGGER IF NOT EXISTS trg_notes_after_delete
                AFTER DELETE ON notes
            BEGIN
                INSERT INTO notes_fts(notes_fts, rowid, key, title,
                                      memo, tags, full_text, summary)
                VALUES ('delete', old.key, old.key, old.title,
                        old.memo, old.tags, old.full_text, old.summary);
            END;

            CREATE TRIGGER IF NOT EXISTS trg_notes_after_update
                AFTER UPDATE ON notes
            BEGIN
                INSERT INTO notes_fts(notes_fts, rowid, key, title,
                                      memo, tags, full_text, summary)
                VALUES ('delete', old.key, old.key, old.title,
                        old.memo, old.tags, old.full_text, old.summary);
                INSERT INTO notes_fts(rowid, key, title,
                                      memo, tags, full_text, summary)
                VALUES (new.key, new.key, new.title,
                        new.memo, new.tags, new.full_text, new.summary);
            END;
        """
            )
            cursor.executescript(trigger_sql)

            # --- リンクテーブル作成 ▼▼▼ ---
            cursor.execute(
                """
            CREATE TABLE IF NOT EXISTS note_links (
                source_key TEXT NOT NULL,
                target_key TEXT NOT NULL,
                PRIMARY KEY (source_key, target_key)
            )
            """
            )

            cursor.execute(
                """
            CREATE INDEX IF NOT EXISTS idx_target_key
                ON note_links (target_key)
            """
            )

            conn.commit()

            # 2. 既存のキーをDBから取得
            existing_keys = set()
            try:
                # 既存キーを取得
                current_keys_df = pd.read_sql_query(
                    f"SELECT key FROM {table_name}", conn
                )
                existing_keys = set(current_keys_df["key"])
            except Exception:
                # テーブルが今作られたばかりならデータはない
                pass

            # 3. 既存キーと重複しないノートのみをフィルタリング
            keys_to_append = df_new_notes["key"]
            df_to_append = df_new_notes[~keys_to_append.isin(existing_keys)]

            if df_to_append.empty:
                logger.info("DBに追記する新規ノートはありません（すべて重複）。")
                return

            # 4. 新規ノートのみをDBに追記
            #    (カラムが完全一致しなくても追記できるよう、必要なカラムを揃える)
            all_columns = [
                "date",
                "time",
                "title",
                "pages",
                "tags",
                "key",
                "memo",
                "commonplace_key",
                "filepath",
                "full_text",
                "merged_pdf_filename",
                "merged_start_page",
                "summary",
            ]

            # 追記用DataFrameのカラムをマスターリストに合わせて整える
            df_final_append = pd.DataFrame(columns=all_columns)
            for col in all_columns:
                if col in df_to_append.columns:
                    df_final_append[col] = df_to_append[col]
                else:
                    df_final_append[col] = ""  # 足りないカラムは空文字で埋める

            df_final_append.to_sql(table_name, conn, if_exists="append", index=False)

            # 5. 追記したノートのリンク情報を解析して note_links に書き込む
            #    (Nexusのutils.pyにあるヘルパーを流用)
            logger.info(
                f"{len(df_to_append)} 件の新規ノートのリンクを解析・登録します..."
            )

            for _, row in df_to_append.iterrows():
                source_key = row.get("key")
                memo_text = row.get("memo", "")
                if source_key and memo_text:
                    # (utils._update_note_links は DELETE & INSERT を行う)
                    NexusUtils._update_note_links(cursor, source_key, memo_text)

            conn.commit()  # リンクテーブルの変更をコミット

            logger.info(
                f"{len(df_to_append)} 件の新規ノートを "
                + f"{master_db_path.name} に追記しました。"
            )

        except Exception as e:
            conn.rollback()
            raise Exception(f"マスターDBへの追記に失敗しました: {e}")
        finally:
            conn.close()

    def save_merged_index_csv(self, notes_with_merged_info, merged_pdf_path):
        csv_filepath = Path(merged_pdf_path).with_suffix(".csv")
        try:
            with open(csv_filepath, "w", newline="", encoding="utf-8-sig") as f:

                header = [
                    "date",
                    "time",
                    "title",
                    "pages",
                    "tags",
                    "key",
                    "memo",
                    "commonplace_key",
                    "filepath",
                    "merged_pdf_filename",
                    "merged_start_page",
                    "summary",
                ]
                writer = csv.DictWriter(f, fieldnames=header, extrasaction="ignore")
                writer.writeheader()
                for note in notes_with_merged_info:
                    note_to_write = note.copy()
                    note_to_write["tags"] = ";".join(
                        sorted(note_to_write.get("tags", []))
                    )
                    writer.writerow(note_to_write)
        except Exception as e:
            messagebox.showerror(
                "CSV保存エラー", f"統合後目次CSVの保存に失敗しました: {e}"
            )

    def scan_folder(self):
        folder_path = tkinter.filedialog.askdirectory(
            title="新規読み込みするフォルダを選択"
        )
        if not folder_path:
            return
        self.label.configure(text=f"読み込み中: {folder_path}")
        self.update_idletasks()
        target_dir = Path(folder_path)
        self.all_notes_info = [
            info
            for pdf_file in target_dir.glob("*.pdf")
            if (info := Process.get_note_info(pdf_file, self.key_rect))
        ]

        # ノート情報に 'summary' がなければ追加（get_note_infoで取得できた場合は維持）
        for info in self.all_notes_info:
            if "summary" not in info:
                info["summary"] = ""

        side_note_suffix = "_Note"

        # 1. 親ノートの「タイトル」と「Index Key」の対応辞書を作成する
        #    (get_note_info が返す 'title' をキーにする)
        parent_key_map = {}
        for info in self.all_notes_info:
            # pdf_processor が抽出した title を取得
            # (例: "20241025_Example" -> "Example", "Example.pdf" -> "Example")
            title = info.get("title", "")
            key = info.get("commonplace_key", "")

            # "_Note" で終わっておらず、かつ Index Key が設定されているノートを親とみなす
            if not title.endswith(side_note_suffix) and key:
                parent_key_map[title] = key

        # 2. もう一度全ノートをスキャンし、サイドノートにKeyを継承させる
        keys_inherited_count = 0
        for info in self.all_notes_info:
            title = info.get("title", "")

            # title が "_Note" で終わり、かつ Index Key が空の場合
            if title.endswith(side_note_suffix) and not info.get("commonplace_key"):

                # 親のタイトル名を取得 (例: "Example_Note" -> "Example")
                parent_title = title[: -len(side_note_suffix)]

                # 親がマップに存在すれば、そのKeyを継承する
                if parent_title in parent_key_map:
                    info["commonplace_key"] = parent_key_map[parent_title]
                    keys_inherited_count += 1

        if keys_inherited_count > 0:
            logger.debug(
                f"{keys_inherited_count}件のサイドノートにIndex Keyを継承しました。"
            )

        self.all_notes_info.sort(key=lambda note: (note["date"], note["time"]))
        self.deselect_all()  # 読み込み時は選択をリセット
        self.update_note_list()
        self.label.configure(
            text=f"読み込み完了！ {len(self.all_notes_info)}件のファイルを読み込みました。"
        )

    def update_note_list(self):
        """
        リストUIを再描画する。
        未登録のIndex Keyがある場合、黄色/オレンジ色で強調表示する。
        [改善] 左クリックで行編集、右クリックで [[Key: Title]] リンクをコピーする。
        """
        for widget in self.scrollable_frame.winfo_children():
            widget.destroy()

        if not self.all_notes_info:
            ctk.CTkLabel(
                self.scrollable_frame, text="PDFファイルが見つかりませんでした。"
            ).pack()
        else:
            default_text_color = ("#1F1F1F", "#1F1F1F")
            warning_text_color = ("#f08300", "#FF4500")
            # 未登録キー用の注意色 (ライトモード: 濃いオレンジ, ダークモード: 黄色)
            unknown_key_text_color = ("#D35400", "#FFC107")

            for note_data in self.all_notes_info:
                row_frame = ctk.CTkFrame(self.scrollable_frame, fg_color="transparent")

                # チェックボックスの追加
                note_key = note_data.get("key")
                var = ctk.StringVar(
                    value=(
                        "on" if note_key and note_key in self.selected_notes else "off"
                    )
                )

                checkbox = ctk.CTkCheckBox(
                    row_frame,
                    text="",
                    variable=var,
                    onvalue="on",
                    offvalue="off",
                    width=25,
                    command=lambda note=note_data, v=var: self.toggle_selection(
                        note, v.get()
                    ),
                )
                # keyがないノート (読み込み失敗など) は選択不可に
                if not note_key:
                    checkbox.configure(state="disabled")

                checkbox.pack(side="left", padx=(0, 5))

                # --- Index Key の判定 ---
                cp_key = note_data.get("commonplace_key", "")

                # config.ini に登録されているキーと一致するか確認
                is_registered_key = cp_key in self.commonplace_key_options

                # アイコン表示 (小文字化して辞書検索)
                icon = self.key_icons.get(cp_key.lower(), "")
                icon_color = self.key_colors.get(cp_key.lower(), default_text_color)

                icon_label = None  # icon_labelを初期化
                if icon:
                    icon_label = ctk.CTkLabel(
                        row_frame, text=icon, text_color=icon_color, font=("", 14)
                    )
                    icon_label.pack(side="left", padx=(0, 5))

                # --- テキストラベルの構築 ---
                key_display = f" [ID: {note_key}]" if note_key else ""
                tags_display = (
                    " [タグ: " + ", ".join(sorted(note_data.get("tags", []))) + "]"
                    if note_data.get("tags")
                    else ""
                )

                display_text = ""
                text_color = default_text_color

                if note_data.get("is_warning"):
                    # ファイル名解析失敗などの重大な警告
                    display_text = (
                        f"【警告】[{note_data.get('date')}] "
                        f"{note_data.get('title')}{key_display}{tags_display}"
                    )
                    text_color = warning_text_color

                elif cp_key and not is_registered_key:
                    # ★未登録のIndex Key (OCR誤認識など) の場合
                    t = note_data.get("time", "")
                    time_display = (
                        f"({t[0:2]}:{t[2:4]}:{t[4:6]})" if t != "999999" else ""
                    )

                    # 黄色文字で「未登録: XXX」と表示
                    display_text = (
                        f"【未登録: {cp_key}】 "
                        f"日付: {note_data.get('date')} {time_display},{key_display} "
                        f"タイトル: {note_data.get('title')}{tags_display}"
                    )
                    text_color = unknown_key_text_color

                else:
                    # 正常なノート
                    t = note_data.get("time", "")
                    time_display = (
                        f"({t[0:2]}:{t[2:4]}:{t[4:6]})" if t != "999999" else ""
                    )
                    display_text = (
                        f"日付: {note_data.get('date')} {time_display},{key_display} "
                        f"タイトル: {note_data.get('title')}{tags_display}"
                    )
                    text_color = default_text_color

                text_label = ctk.CTkLabel(
                    row_frame, text=display_text, text_color=text_color, anchor="w"
                )
                text_label.pack(side="left")

                # --- イベントバインド ---

                # 1. 左クリック (編集ウィンドウを開く)
                def edit_command(e, note=note_data):
                    self.open_data_editor(note)

                row_frame.bind("<Button-1>", edit_command)  # フレーム全体
                text_label.bind("<Button-1>", edit_command)
                if icon_label:
                    icon_label.bind("<Button-1>", edit_command)

                # 2. 右クリック ([[Key: Title]] リンクをコピー)
                if note_key:
                    # [変更] クロージャで note_data 全体をキャプチャ
                    def create_copy_key_handler(note_to_copy):
                        def handler(event):
                            try:
                                # [変更] KeyとTitleを取得
                                key = note_to_copy.get("key", "")
                                title = note_to_copy.get("title", "")

                                # [変更] [[Key: Title]] 形式の文字列を生成
                                text_to_copy = f"[[{key}: {title}]]"

                                self.clipboard_clear()
                                self.clipboard_append(text_to_copy)  # 変更後の文字列
                                self.update()  # クリップボードを確定
                                logger.info(
                                    f"リンクをクリップボードにコピーしました: {text_to_copy}",
                                    extra={"sensitive": True},
                                )

                                # (フィードバック)
                                self.label.configure(
                                    text=f"コピーしました: {text_to_copy}"
                                )
                                self.after(2000, lambda: self.label.configure(text=""))
                            except Exception as e:
                                logger.error(f"リンクのコピーに失敗: {e}")

                        return handler

                    copy_command = create_copy_key_handler(note_data)
                    row_frame.bind("<Button-3>", copy_command)
                    text_label.bind("<Button-3>", copy_command)
                    if icon_label:
                        icon_label.bind("<Button-3>", copy_command)

                row_frame.pack(fill="x", padx=5, pady=2)

    def _copy_bookmarks_recursive(self, outline_items, writer, reader, parent=None):
        """
        pypdfの目次(outline)の階層構造を再帰的にたどり、writerにコピーする関数
        """
        i = 0
        while i < len(outline_items):
            item = outline_items[i]

            # 現在のアイテムをブックマークとして追加
            # get_destination_page_numberでページ番号を安全に取得
            page_num = reader.get_destination_page_number(item)
            if page_num is not None:
                new_parent = writer.add_outline_item(
                    item.title, page_num, parent=parent
                )

                # 次の要素がリスト（＝子要素のリスト）かチェック
                if i + 1 < len(outline_items) and isinstance(
                    outline_items[i + 1], list
                ):
                    # 子要素のリストに対して再帰的にこの関数を呼び出す
                    self._copy_bookmarks_recursive(
                        outline_items[i + 1], writer, reader, parent=new_parent
                    )
                    i += 1  # 子要素リストをスキップするためインデックスを1つ進める
            i += 1

    def open_data_editor(self, note_data):
        session_tags = set()
        for note in self.all_notes_info:
            session_tags.update(note.get("tags", []))
        combined_tags = session_tags.union(set(self.predefined_tags))
        if hasattr(self, "editor_window") and self.editor_window.winfo_exists():
            self.editor_window.focus()
            return
        self.editor_window = Dialogs.DataEditorWindow(
            self, note_data, list(combined_tags), self.commonplace_key_options
        )

    def generate_pdf(self):
        # メタデータ作成用にdatetimeが必要
        from datetime import datetime

        if not self.all_notes_info:
            self.label.configure(text="PDF生成対象のデータがありません。")
            return

        # ---------------------------------------------------------
        # 0. 事前フィルタリング (ファイル完全検証)
        # ---------------------------------------------------------
        self.label.configure(text="PDF生成準備中: ファイル検証とテキスト抽出...")
        self.update_idletasks()

        valid_notes_info = []
        excluded_count = 0

        # 検証ループ
        for note in self.all_notes_info:
            # パス解決
            path_str = note.get("filepath", "")
            if not path_str:
                excluded_count += 1
                continue
            pdf_path = Path(path_str)

            # A. ファイル存在チェック
            if not pdf_path.is_file():
                logger.warning(f"除外(未検出): {pdf_path.name}")
                excluded_count += 1
                continue

            # B. PDF有効性 & ページ数チェック
            try:
                # 実際に開いてページ数を確認 (0ページや破損を弾く)
                reader = PdfReader(str(pdf_path))
                real_pages = len(reader.pages)

                if real_pages <= 0:
                    logger.warning(f"除外(ページなし): {pdf_path.name}")
                    excluded_count += 1
                    continue

                # メモリ上のページ数を実態に合わせて更新
                note["pages"] = real_pages

            except Exception as e:
                logger.warning(f"除外(読込不可): {pdf_path.name} - {e}")
                excluded_count += 1
                continue

            # C. タイトルの正規化 (No Title 対応)
            # pdf_processor.py で "" が返ってくるようになったため、ここで "No Title" をセット
            title = note.get("title")
            if not title or not str(title).strip():
                note["title"] = "No Title"

            # D. 本文テキスト(full_text)抽出
            if not note.get("full_text"):
                try:
                    extracted_text = Process.get_full_text(pdf_path)
                    note["full_text"] = extracted_text
                except Exception as e:
                    logger.warning(
                        f"{pdf_path.name} のfull_text抽出に失敗: {e}",
                        extra={"sensitive": True},
                    )

            # 全チェック通過
            valid_notes_info.append(note)

        # 結果通知
        if excluded_count > 0:
            msg = f"{excluded_count} 件のファイルが無効（存在しない・破損）なため、生成対象から除外しました。\nログを確認してください。"
            logger.warning(msg)
            messagebox.showwarning("ファイル除外", msg)

        if not valid_notes_info:
            messagebox.showerror(
                "エラー", "有効なPDFファイルが1つもありません。処理を中断します。"
            )
            self.label.configure(text="有効なファイルがありません。")
            return

        # ---------------------------------------------------------
        # 1. 日付と保存先の入力
        # ---------------------------------------------------------
        dialog = Dialogs.DateInputDialog(self)
        input_result = dialog.get_input()

        if not input_result:
            return

        # 戻り値を3つ受け取る (年, 月, 画像パス)
        year, month, cover_image_path = input_result

        base_pdf_title = f"{self.reportlab_title_prefix} ({year}年 {month}月)"

        save_filepath_str = tkinter.filedialog.asksaveasfilename(
            defaultextension=".pdf",
            filetypes=[("PDFファイル", "*.pdf")],
            title="統合PDFの保存先を選択 (分割時は連番が付与されます)",
            initialfile=f"統合ノート_{year}_{month:02d}.pdf",
        )
        if not save_filepath_str:
            return

        base_save_path = Path(save_filepath_str)

        # ---------------------------------------------------------
        # 2. ノートの分割
        # ---------------------------------------------------------
        volumes = self._split_notes_into_volumes(
            valid_notes_info, self.max_pages_per_volume
        )
        total_volumes = len(volumes)

        if total_volumes > 1:
            msg = (
                f"指定されたページ制限 ({self.max_pages_per_volume} p) に基づき、"
                f"全体を {total_volumes} 冊に分割して生成します。"
            )
            logger.info(msg)
            messagebox.showinfo("分割生成", msg)

        # ======================================================================
        # 各巻ごとのループ処理
        # ======================================================================
        for vol_idx, notes_in_volume in enumerate(volumes):
            current_vol_num = vol_idx + 1

            if total_volumes > 1:
                current_pdf_title = f"{base_pdf_title} Vol.{current_vol_num}"
                stem = base_save_path.stem
                current_save_path = base_save_path.with_name(
                    f"{stem}_Vol{current_vol_num}{base_save_path.suffix}"
                )
            else:
                current_pdf_title = base_pdf_title
                current_save_path = base_save_path

            self.label.configure(
                text=f"[{current_vol_num}/{total_volumes}冊目] PDF生成中..."
            )
            self.update_idletasks()

            # --- 一時ディレクトリの作成 ---
            temp_dir = tempfile.mkdtemp()
            try:
                # --- A. 骨格PDF生成 ---
                self.label.configure(
                    text=f"[{current_vol_num}/{total_volumes}] ページ構成を生成中 (ReportLab)..."
                )
                self.update_idletasks()

                gen_config = {
                    "font_path": self.font_path,
                    "reportlab_font": self.reportlab_font,
                    "reportlab_author": self.reportlab_author,
                    "key_icons": self.key_icons,
                    "key_colors": self.key_colors,
                }

                pdf_gen = Generator.ReportLabPDFGenerator(gen_config)
                draft_pdf_path = Path(temp_dir) / "mokuji.pdf"

                # 骨格生成
                layout_info = pdf_gen.create_skeleton_pdf(
                    notes_in_volume,
                    current_pdf_title,
                    self.paper_size,
                    str(draft_pdf_path),
                    cover_image_path=cover_image_path,
                )

                if not draft_pdf_path.is_file():
                    messagebox.showerror("エラー", "骨格PDFの生成に失敗しました。")
                    return

                # --- B. PDF結合処理 ---
                self.label.configure(
                    text=f"[{current_vol_num}/{total_volumes}] ノートを結合中..."
                )
                self.update_idletasks()

                draft_reader = PdfReader(str(draft_pdf_path))
                final_writer = PdfWriter()

                note_content_start_page = layout_info["content_start_page"]
                index_start_page = layout_info["index_start_page"]

                # 1. 目次
                for i in range(note_content_start_page):
                    final_writer.add_page(draft_reader.pages[i])

                # 2. 本文
                updated_notes_info = []
                note_page_cursor = note_content_start_page

                for note in notes_in_volume:
                    # マージ後の情報を記録 (DB登録用)
                    note["merged_start_page"] = note_page_cursor + 1
                    note["merged_pdf_filename"] = current_save_path.name
                    updated_notes_info.append(note)

                    try:
                        original_reader = PdfReader(note["filepath"])
                        for i in range(len(original_reader.pages)):
                            if note_page_cursor >= index_start_page:
                                break

                            template_page = draft_reader.pages[note_page_cursor]
                            content_page = original_reader.pages[i]
                            content_page.merge_page(template_page)

                            final_writer.add_page(content_page)
                            note_page_cursor += 1
                    except Exception as e:
                        logger.error(f"ノート結合エラー ({note.get('title')}): {e}")
                        p_count = note.get("pages", 0)
                        for _ in range(p_count):
                            if note_page_cursor < len(draft_reader.pages):
                                final_writer.add_page(
                                    draft_reader.pages[note_page_cursor]
                                )
                                note_page_cursor += 1

                # 3. 索引
                for i in range(index_start_page, len(draft_reader.pages)):
                    final_writer.add_page(draft_reader.pages[i])

                # --- C. メタデータ設定---
                # PDFプロパティにタイトルや作成者を書き込む
                metadata = {
                    "/Title": current_pdf_title,
                    "/Author": self.reportlab_author,
                    "/Producer": "Synapsen Ersteller (ReportLab + pypdf)",
                    "/Creator": "Synapsen",
                    "/CreationDate": datetime.now().strftime("D:%Y%m%d%H%M%S"),
                }
                final_writer.add_metadata(metadata)

                # --- D. ブックマークのコピー ---
                if draft_reader.outline:
                    self._copy_bookmarks_recursive(
                        draft_reader.outline, final_writer, draft_reader
                    )

                # --- E. 復旧用メタデータ(JSON)埋め込み ---
                try:
                    metadata_to_embed = []
                    for note in updated_notes_info:
                        clean_note = {
                            "key": note.get("key"),
                            "title": note.get("title"),
                            "date": note.get("date"),
                            "time": note.get("time"),
                            "tags": note.get("tags"),
                            "memo": note.get("memo"),
                            "commonplace_key": note.get("commonplace_key"),
                            "pages": note.get("pages"),
                            "merged_start_page": note.get("merged_start_page"),
                            "merged_pdf_filename": note.get("merged_pdf_filename"),
                            "summary": note.get("summary"),
                        }
                        metadata_to_embed.append(clean_note)

                    data_to_save = {
                        "schema_version": CURRET_SCHEMA_VERSION,
                        "volume_info": f"{current_vol_num}/{total_volumes}",
                        "notes_data": metadata_to_embed,
                    }

                    json_data = json.dumps(data_to_save, ensure_ascii=False, indent=2)
                    final_writer.add_attachment(
                        "synapsen_metadata_backup.json", json_data.encode("utf-8")
                    )
                except Exception as e:
                    logger.warning(f"復旧用メタデータ埋め込み失敗: {e}")

                # --- F. ファイル保存 ---
                self.label.configure(
                    text=f"[{current_vol_num}/{total_volumes}] 保存中..."
                )
                self.update_idletasks()

                with open(current_save_path, "wb") as f:
                    final_writer.write(f)

                # --- G. DB / CSV 保存 ---
                if self.auto_append_db and self.default_db_path:
                    self.append_to_master_db(updated_notes_info)

                if self.create_individual_csv:
                    self.save_merged_index_csv(updated_notes_info, current_save_path)

            except Exception as e:
                logger.error(f"Fatal error in Vol {current_vol_num}: {e}")
                messagebox.showerror(
                    "エラー",
                    f"{current_vol_num}冊目の処理中にエラーが発生しました:\n{e}",
                )
                return
            finally:
                shutil.rmtree(temp_dir)

        self.label.configure(text="全てのPDF生成と保存が完了しました。")
        messagebox.showinfo(
            "完了", f"合計 {total_volumes} 冊のPDFを作成・保存しました。"
        )

    def _split_notes_into_volumes(self, all_notes, max_pages):
        """
        ノートリストを最大ページ数に基づいて分割するヘルパーメソッド。
        max_pages が 0 以下の場合は分割せずそのまま返す。
        """
        if max_pages <= 0:
            return [all_notes]

        volumes = []
        current_volume = []
        current_pages = 0

        # 目次や索引のために、安全マージンとして少しページ数を差し引いて計算しても良いですが、
        # ここでは単純にノートの合計ページ数で判定します。

        for note in all_notes:
            # 実際のファイルが存在しない場合などはページ数0として扱う
            p = note.get("pages", 0)

            # 「現在のページ数 + このノートのページ数」が上限を超える場合、かつ
            # 「現在の巻に既にノートが入っている」場合（1つのノートだけで上限超えなら単独で入れる）
            if current_volume and (current_pages + p > max_pages):
                volumes.append(current_volume)
                current_volume = []
                current_pages = 0

            current_volume.append(note)
            current_pages += p

        # 最後の巻を追加
        if current_volume:
            volumes.append(current_volume)

        return volumes

    def sync_with_folder(self):
        if not self.all_notes_info:
            self.label.configure(text="先にCSVを読み込んでください。")
            return
        folder_path = tkinter.filedialog.askdirectory(title="同期するフォルダを選択")
        if not folder_path:
            return
        self.label.configure(text=f"同期中: {folder_path}")
        self.update_idletasks()
        app_paths = {note.get("filepath") for note in self.all_notes_info}
        disk_paths = {str(pdf_file) for pdf_file in Path(folder_path).glob("*.pdf")}
        added_paths = disk_paths - app_paths
        deleted_paths = app_paths - disk_paths
        added_count, deleted_count = 0, 0
        if deleted_paths:
            deleted_filenames = "\n".join([f"- {Path(p).name}" for p in deleted_paths])
            user_response = messagebox.askyesno(
                "削除の確認",
                f"以下のファイルがフォルダから見つかりませんでした。リストから削除しますか？\n\n{deleted_filenames}",
            )
            if user_response:
                self.all_notes_info = [
                    note
                    for note in self.all_notes_info
                    if note.get("filepath") not in deleted_paths
                ]
                deleted_count = len(deleted_paths)
                # 削除されたノートが選択されていた場合、選択状態からも削除
                keys_in_memory = {note.get("key") for note in self.all_notes_info}
                self.selected_notes = self.selected_notes.intersection(keys_in_memory)

        if added_paths:
            for path in sorted(list(added_paths)):
                info = Process.get_note_info(Path(path), self.key_rect)
                if info:
                    self.all_notes_info.append(info)
            added_count = len(added_paths)

        if added_count > 0 or deleted_count > 0:
            self.all_notes_info.sort(key=lambda note: (note["date"], note["time"]))
            self.update_note_list()  # UIを再描画
            self.update_batch_buttons_state()  # ボタン状態を更新
            self.label.configure(
                text=f"同期完了！ {added_count}件追加, {deleted_count}件削除"
            )
        else:
            self.label.configure(text="変更はありませんでした。")

    def toggle_selection(self, note_data, state_str):
        """
        チェックボックスがクリックされたときに呼び出され、
        ノートの選択状態を切り替える。
        """
        key = note_data.get("key")
        if not key:
            return  # key がないノートは選択不可

        if state_str == "on":
            self.selected_notes.add(key)
        else:
            self.selected_notes.discard(
                key
            )  # removeと違い、存在しなくてもエラーにならない

        self.update_batch_buttons_state()

    def update_batch_buttons_state(self):
        """
        選択されているノートの数に応じて、
        一括編集ボタンと選択解除ボタンの状態を更新する。
        """
        count = len(self.selected_notes)
        if count > 0:
            self.batch_edit_button.configure(text=f"一括編集 ({count})", state="normal")
            self.deselect_all_button.configure(state="normal")
            self.copy_links_button.configure(
                text=f"リンクコピー ({count})", state="normal"
            )
        else:
            self.batch_edit_button.configure(text="一括編集 (0)", state="disabled")
            self.deselect_all_button.configure(state="disabled")
            self.copy_links_button.configure(text="リンクコピー (0)", state="disabled")

    def deselect_all(self):
        """
        すべてのノートの選択を解除し、UIを更新する。
        """
        self.selected_notes.clear()
        self.update_batch_buttons_state()
        self.update_note_list()  # チェックボックスをすべて外すために再描画

    def open_batch_editor(self):
        """
        「一括編集」ボタンが押されたときに、
        一括編集ダイアログ (BatchEditWindow) を開く。
        """
        if not self.selected_notes:
            return

        # ダイアログの「既存タグから選択」で使うためのタグリストを作成
        session_tags = set()
        for note in self.all_notes_info:
            session_tags.update(note.get("tags", []))
        combined_tags = session_tags.union(set(self.predefined_tags))

        # gui_dialogs.py に追加した BatchEditWindow を呼び出す
        dialog = Dialogs.BatchEditWindow(
            self,
            len(self.selected_notes),
            list(combined_tags),
            self.commonplace_key_options,
        )

        # ダイアログの結果待機
        result = dialog.get_input()

        if result:
            # ユーザーが「適用」を押した場合
            self.apply_batch_edits(
                result.get("index_key"),
                result.get("tags_to_add", []),
                result.get("tags_to_remove", []),
            )

    def apply_batch_edits(self, index_key_to_set, tags_to_add, tags_to_remove):
        """
        BatchEditWindow で指定された内容に基づき、
        選択中のすべてのノートのデータを変更する。
        """
        if index_key_to_set is None and not tags_to_add and not tags_to_remove:
            logger.info("一括編集: 変更内容がありません。")
            return

        modified_count = 0

        # メモリ上の self.all_notes_info を直接変更
        for note in self.all_notes_info:
            note_key = note.get("key")
            if note_key and note_key in self.selected_notes:

                # 1. Index Key の設定
                if index_key_to_set is not None:
                    note["commonplace_key"] = index_key_to_set

                # 2. タグの変更
                current_tags = set(note.get("tags", []))

                # 2a. タグの追加 (階層も考慮)
                for tag_to_add in tags_to_add:
                    parts = tag_to_add.split("_")
                    for i in range(len(parts)):
                        hierarchical_tag = "_".join(parts[: i + 1])
                        current_tags.add(hierarchical_tag)

                # 2b. タグの削除
                current_tags.difference_update(tags_to_remove)

                note["tags"] = sorted(list(current_tags))
                modified_count += 1

        logger.info(f"{modified_count} 件のノートを一括編集しました。")

        # 変更をUIに反映
        self.deselect_all()  # 選択解除 (UI再描画も含まれる)

    def copy_selected_links(self):
        """
        [新規]
        選択されたノートのリンク文字列（[[Key: Title]]）を生成し、
        クリップボードにコピーする。
        (Synapsen_Nexus_main.py の同名メソッドを Ersteller 用に移植)
        """
        if not self.selected_notes:
            self.label.configure(text="コピー対象のノートが選択されていません。")
            return

        if not self.all_notes_info:
            return

        # 1. 選択されたノートの辞書を抽出
        selected_notes_data = []
        for note in self.all_notes_info:
            key = note.get("key")
            if key and key in self.selected_notes:
                selected_notes_data.append(note)

        if not selected_notes_data:
            return

        # 2. 日付・時刻順にソート (Erstellerのデフォルト順)
        selected_notes_data.sort(key=lambda note: (note["date"], note["time"]))

        # 3. リンク文字列を生成
        link_texts = []
        for note in selected_notes_data:
            key = note["key"]
            title = note["title"]
            # Synapsenのリンク形式: [[Key: Title]]
            link_texts.append(f"[[{key}: {title}]]")

        # 4. 改行区切りで結合
        clipboard_text = "\n".join(link_texts)

        # 5. クリップボードへコピー
        try:
            self.clipboard_clear()
            self.clipboard_append(clipboard_text)
            self.update()  # クリップボード更新を確定

            messagebox.showinfo(
                "コピー完了",
                f"{len(link_texts)}件のリンクをクリップボードにコピーしました。\n"
                "ノートのメモ欄にペーストして利用できます。",
                parent=self,
            )
            self.label.configure(text=f"{len(link_texts)}件のリンクをコピーしました。")

        except Exception as e:
            logger.error(f"リンクのコピーに失敗: {e}")
            messagebox.showerror(
                "コピー失敗",
                f"クリップボードへのコピーに失敗しました:\n{e}",
                parent=self,
            )

    def open_recovery_tool(self):
        """DB復旧ツールウィンドウを開く"""
        # 現在設定されているDBパスをデフォルトとして渡す
        default_db = self.default_db_path if self.default_db_path else ""

        # ツールウィンドウの作成
        recovery_win = db_recovery_tool.DBRecoveryWindow(
            self, default_db_path=default_db
        )
        recovery_win.focus()

    def open_config_editor(self):
        """設定編集ウィンドウを開く"""
        try:
            # config_editor.py が同じフォルダにあることを前提にインポート
            from config_editor import ConfigEditorWindow
        except ImportError:
            messagebox.showerror(
                "エラー",
                "config_editor.py が見つかりません。\nSynapsen_Ersteller フォルダに配置してください。",
            )
            return

        # config.ini のパスを解決 (load_config と同じロジック)
        if getattr(sys, "frozen", False):
            # .exe実行の場合
            base_path = os.path.dirname(sys.executable)
            config_path = os.path.join(base_path, "config.ini")
        else:
            # スクリプト実行の場合
            base_path = os.path.dirname(os.path.abspath(__file__))
            config_path = os.path.join(
                os.path.abspath(os.path.join(base_path, "..")), "config.ini"
            )

        # ウィンドウを開く
        ConfigEditorWindow(self, config_path)


if __name__ == "__main__":
    app = Synapsen_Ersteller()

    if app.icon_path:  # <-- クラス内で取得したパスを利用
        try:
            # 'default=' を指定し、OSダイアログ(エクスプローラ等)にも適用
            app.iconbitmap(default=str(app.icon_path))
        except Exception as e:
            logger.error(f"Icon default setting error: {e}")
    else:
        logger.warning(
            "警告: アイコンファイル (assets/synapsen.ico) が見つかりません。"
        )

    app.mainloop()

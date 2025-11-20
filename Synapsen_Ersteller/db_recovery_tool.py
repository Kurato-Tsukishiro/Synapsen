# === 1. 標準ライブラリ ===
import sys
from pathlib import Path
import json
import unicodedata
from textwrap import dedent
import logging

# === 2. プロジェクトルートをパスに追加 (ここからE402の原因) ===
current_dir = Path(__file__).parent
root_dir = current_dir.parent
if str(root_dir) not in sys.path:
    sys.path.append(str(root_dir))

# === 3. プロジェクト内モジュールとサードパーティ (E402を抑制) ===
import customtkinter as ctk                     # noqa: E402
from tkinter import filedialog, messagebox      # noqa: E402
import pandas as pd                             # noqa: E402
import sqlite3                                  # noqa: E402
from pypdf import PdfReader                     # noqa: E402
import fitz                                     # noqa: E402
from Synapsen_Nexus import utils as NexusUtils  # noqa: E402

logger = logging.getLogger(__name__)


class DBRecoveryWindow(ctk.CTkToplevel):
    """
    統合PDFに埋め込まれたメタデータ(JSON)から、
    マスターDBを復旧・追記するためのウィンドウ。
    """
    def __init__(self, parent, default_db_path=None):
        super().__init__(parent)
        self.parent = parent
        self.title("DB復旧・再構築ツール")
        self.geometry("500x450")

        # アイコン設定 (親から継承)
        if hasattr(parent, 'icon_path') and parent.icon_path:
            try:
                self.iconbitmap(default=str(parent.icon_path))
            except Exception:
                pass

        self.grab_set()  # モーダル化

        # --- UI配置 ---
        self.grid_columnconfigure(1, weight=1)

        # 1. 説明ラベル
        ctk.CTkLabel(
            self,
            text="統合PDF内のバックアップデータから、DBを復旧します。\n"
                 "※ 元のPDFを削除していても、ここからメタデータと本文を復元できます。",
            justify="left", text_color="gray"
        ).grid(row=0, column=0, columnspan=3, padx=10, pady=10, sticky="w")

        # 2. ソースPDF選択
        ctk.CTkLabel(
            self, text="ソースPDF:"
        ).grid(row=1, column=0, padx=10, pady=5, sticky="w")
        self.pdf_path_entry = ctk.CTkEntry(
            self, placeholder_text="メタデータ入りの統合PDFを選択...")
        self.pdf_path_entry.grid(row=1, column=1, padx=5, pady=5, sticky="ew")
        ctk.CTkButton(
            self, text="参照", width=60, command=self.browse_pdf
        ).grid(row=1, column=2, padx=10, pady=5)

        # 3. ターゲットDB選択
        ctk.CTkLabel(
            self, text="復旧先DB:"
        ).grid(row=2, column=0, padx=10, pady=5, sticky="w")
        self.db_path_entry = ctk.CTkEntry(
            self, placeholder_text="Synapsen_Master.db")
        self.db_path_entry.grid(row=2, column=1, padx=5, pady=5, sticky="ew")

        if default_db_path:
            self.db_path_entry.insert(0, str(default_db_path))

        ctk.CTkButton(
            self, text="参照", width=60, command=self.browse_db
        ).grid(row=2, column=2, padx=10, pady=5)

        # 4. 情報表示エリア
        self.info_textbox = ctk.CTkTextbox(self, height=150)
        self.info_textbox.grid(
            row=3, column=0, columnspan=3, padx=10, pady=10, sticky="nsew")
        self.info_textbox.insert("1.0", "PDFを選択して「内容確認」を押してください。\n")
        self.info_textbox.configure(state="disabled")

        # 5. ボタンエリア
        btn_frame = ctk.CTkFrame(self, fg_color="transparent")
        btn_frame.grid(row=4, column=0, columnspan=3, pady=10)

        ctk.CTkButton(
            btn_frame, text="1. 内容確認 (スキャン)", command=self.scan_pdf
        ).pack(side="left", padx=10)

        self.restore_button = ctk.CTkButton(
            btn_frame, text="2. DBに復元 (実行)",
            command=self.execute_restore, state="disabled",
            fg_color="#D9534F", hover_color="#C9302C"  # 注意を促す赤系
        )
        self.restore_button.pack(side="left", padx=10)

        # 内部保持用データ
        self.extracted_data = []

    def log(self, message):
        self.info_textbox.configure(state="normal")
        self.info_textbox.insert("end", message + "\n")
        self.info_textbox.see("end")
        self.info_textbox.configure(state="disabled")

    def browse_pdf(self):
        path = filedialog.askopenfilename(
            filetypes=[("PDF Files", "*.pdf")], title="統合PDFを選択"
        )
        if path:
            self.pdf_path_entry.delete(0, "end")
            self.pdf_path_entry.insert(0, path)
            self.restore_button.configure(state="disabled")
            self.extracted_data = []

    def browse_db(self):
        path = filedialog.asksaveasfilename(
            filetypes=[("SQLite DB", "*.db")],
            title="復旧先DBを選択 (新規作成または上書き)",
            initialfile="Synapsen_Master_Restored.db"
        )
        if path:
            self.db_path_entry.delete(0, "end")
            self.db_path_entry.insert(0, path)

    def scan_pdf(self):
        """PDF内の添付ファイルをチェックし、データをメモリに展開する"""
        pdf_path = self.pdf_path_entry.get()
        if not pdf_path or not Path(pdf_path).is_file():
            messagebox.showerror("エラー", "有効なPDFファイルを選択してください。")
            return

        self.log("-" * 30)
        self.log(f"スキャン中: {Path(pdf_path).name} ...")
        self.update_idletasks()

        try:
            reader = PdfReader(pdf_path)
            attachments = reader.attachments

            target_filename = "synapsen_metadata_backup.json"

            if not attachments or target_filename not in attachments:
                self.log(f"[エラー] 復旧用データ ({target_filename}) が見つかりません。")
                self.log("このPDFは新しいバージョンのSynapsenで作成されたものではない可能性があります。")
                return

            # データの読み出し
            json_bytes = attachments[target_filename][0]

            # スキーマバージョンを考慮してJSONをパース
            json_data = json.loads(json_bytes.decode('utf-8'))

            schema_ver = 1.0
            notes_list = []

            if isinstance(json_data, dict):
                schema_ver = json_data.get("schema_version", 1.0)
                notes_list = json_data.get("notes_data", [])
                self.log(f"スキーマバージョン: {schema_ver} (新形式) を検出。")
            elif isinstance(json_data, list):
                schema_ver = 1.0  # スキーマを実装する前のノートはスキーマ1.0とみなす
                notes_list = json_data
                self.log("スキーマバージョン: 1.0 として読み込みます。")

            # 抽出したノートリストを内部に保持
            self.extracted_data = notes_list

            for note in self.extracted_data:
                # 'summary' キーがない場合、空文字列で初期化
                if 'summary' not in note:
                    note['summary'] = ""

            count = len(self.extracted_data)
            self.log(f"[成功] {count} 件のノートデータが見つかりました。")

            if count > 0:
                sample = self.extracted_data[0]
                self.log(f"データ例: {sample.get('date')} - {sample.get('title')}")
                self.restore_button.configure(state="normal")
            else:
                self.log("データ件数が0件です。")

        except Exception as e:
            self.log(f"[例外] スキャン中にエラー: {e}")
            logger.error(f"Scan Error: {e}")

    def execute_restore(self):
        """メモリ上のデータをDBにINSERTする"""
        db_path = self.db_path_entry.get()
        pdf_path = self.pdf_path_entry.get()

        if not db_path:
            messagebox.showerror("エラー", "復旧先のデータベースパスを指定してください。")
            return

        if not self.extracted_data:
            return

        # 確認ダイアログ
        ans = messagebox.askyesno(
            "実行確認",
            f"{len(self.extracted_data)} 件のデータを以下のDBに復元します。\n\n"
            f"復元先: {Path(db_path).name}\n\n"
            "既存のデータがある場合、Keyが重複するノートは無視(スキップ)されます。\n"
            "実行しますか？"
        )
        if not ans:
            return

        self.log("-" * 30)
        self.log("DB復元を開始します...")
        self.update_idletasks()

        # ソースPDFのファイル名を取得
        current_pdf_name = Path(pdf_path).name

        # 統合PDFから本文テキスト(full_text)を再抽出するためのPDFドキュメントを開く
        doc = None
        conn = None  # 1. conn を None で初期化

        try:
            doc = fitz.open(pdf_path)
        except Exception as e:
            self.log(f"[エラー] テキスト抽出のためにPDFを開けませんでした: {e}")
            return  # doc が None のまま finally に進み、安全に終了

        try:
            # DataFrameに変換
            df = pd.DataFrame(self.extracted_data)

            # 統合PDFファイル名の上書き
            df['merged_pdf_filename'] = current_pdf_name

            # タグリストを文字列(;)に戻す
            if 'tags' in df.columns:
                df['tags'] = df['tags'].apply(
                    lambda x: ";".join(sorted(x)) if isinstance(x, list) else x
                )

            self.log("PDFから本文テキストを再抽出しています...")
            self.update_idletasks()

            def extract_text_from_range(row):
                """統合PDFから本文テキスト(full_text)を再抽出する

                Args:
                    row (pd.Series): DataFrameの行データ (ノート情報)

                Returns:
                    str: 抽出されたテキスト
                """
                try:
                    # ページ情報の取得
                    start_page_1based = int(row.get('merged_start_page', 0))
                    page_count = int(row.get('pages', 0))

                    if start_page_1based < 1 or page_count < 1:
                        return ""

                    # 0-indexedに変換
                    start_idx = start_page_1based - 1
                    end_idx = start_idx + page_count

                    # 範囲チェック
                    if start_idx >= len(doc):
                        return ""

                    # 指定範囲のページからテキストを結合
                    full_text = ""

                    # end_idx は len(doc) を超えないように制限
                    actual_end = min(end_idx, len(doc))

                    for i in range(start_idx, actual_end):
                        full_text += doc[i].get_text() + "\n"

                    # Unicode正規化 (NFKC) を実行
                    if full_text:
                        normalized_text = (
                            unicodedata.normalize('NFKC', full_text)
                        )
                        return normalized_text.strip()

                    return ""

                except Exception as e:
                    logger.error(
                        f"Text extraction error for {row.get('key')}: {e}"
                    )
                    return ""

            # 各行に対してテキスト抽出を実行
            df['full_text'] = df.apply(extract_text_from_range, axis=1)

            # 復元データ (JSON) が 'memo' を持っていることを確認
            if 'memo' not in df.columns:
                df['memo'] = ""  # 万が一ない場合は空文字で埋める

            conn = sqlite3.connect(db_path)  # 2. conn に代入
            cursor = conn.cursor()

            # 1. 'notes' テーブル
            #    'key' を PRIMARY KEY に明示 (FTS連携に重要)
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

            # 2. 'notes_fts' FTS5 仮想テーブル (新規)
            #    'key' を content_rowid として 'notes' テーブルと紐付け
            create_fts_sql = """
            CREATE VIRTUAL TABLE IF NOT EXISTS notes_fts USING fts5(
                key,
                title,
                memo,
                tags,
                full_text,
                summary,
                content='notes',
                content_rowid='key'
            );
            """
            cursor.execute(create_fts_sql)

            # 3. 自動同期トリガー (新規)
            #    'notes' に変更があったら 'notes_fts' も自動更新する
            trigger_sql = dedent("""
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
        """)
            cursor.executescript(trigger_sql)
            cursor.execute("""
            CREATE TABLE IF NOT EXISTS note_links (
                source_key TEXT NOT NULL,
                target_key TEXT NOT NULL,
                PRIMARY KEY (source_key, target_key)
            )
            """)
            cursor.execute("""
            CREATE INDEX IF NOT EXISTS idx_target_key
                ON note_links (target_key)
            """)

            conn.commit()

            # 重複チェックとインサート
            existing_keys = set()
            try:
                existing = pd.read_sql("SELECT key FROM notes", conn)
                existing_keys = set(existing['key'].astype(str))
            except Exception:
                pass

            df_to_insert = df[~df['key'].isin(existing_keys)]

            skipped_count = len(df) - len(df_to_insert)
            insert_count = len(df_to_insert)

            if insert_count > 0:
                df_to_insert.to_sql(
                    'notes', conn, if_exists='append', index=False)

                # リンクテーブル書き込み
                logger.info(f"{len(df_to_insert)} 件の復元ノートのリンクを解析・登録します...")
                for _, row in df_to_insert.iterrows():
                    source_key = row.get("key")
                    memo_text = row.get("memo", "")  # 復元データ(JSON)のmemoを使用
                    if source_key and memo_text:
                        NexusUtils._update_note_links(
                            cursor, source_key, memo_text
                        )

                conn.commit()  # リンクテーブルの変更をコミット
                self.log(f"[完了] {insert_count} 件を追記しました。")
            else:
                self.log("[完了] 新規データはありませんでした (すべて重複)。")

            if skipped_count > 0:
                self.log(f"(重複のためスキップ: {skipped_count} 件)")

            messagebox.showinfo(
                "完了", "データベースの復元が完了しました。\n(本文テキストもunicode正規化されました)")
            self.destroy()

        except Exception as e:
            if conn:
                conn.rollback()
            self.log(f"[エラー] DB書き込み中にエラー: {e}")
            messagebox.showerror("エラー", f"復元に失敗しました:\n{e}")

        finally:
            # doc と conn を両方ともここで閉じる
            if doc:
                doc.close()
            if conn:
                conn.close()

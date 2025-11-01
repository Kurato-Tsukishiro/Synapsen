import os
import tkinter
import csv
import sys
from tkinter import messagebox
from pathlib import Path
from pypdf import PdfReader, PdfWriter, Transformation
import customtkinter as ctk
import subprocess
import shutil
import tempfile
import configparser

import PDFMargeHelper as Helper
import pdf_processor as Process
import latex_generator as Generator
import gui_dialogs as Dialogs


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
        self.grid_rowconfigure(2, weight=1)

        # 用紙サイズを保持する変数
        self.paper_size = "A4"  # デフォルト
        self.paper_width = Helper.A4_WIDTH
        self.paper_height = Helper.A4_HEIGHT

        self.load_config()

        self.all_notes_info = []
        self.predefined_tags = []
        self.load_predefined_tags()

        self.label = ctk.CTkLabel(
            self, text="Synapsen Normalisiererで処理済みのフォルダを読み込んでください。"
            )
        self.label.grid(row=0, column=0, padx=10, pady=10, sticky="ew")

        top_button_frame = ctk.CTkFrame(self, fg_color="transparent")
        top_button_frame.grid(row=1, column=0, padx=10, pady=5, sticky="ew")
        ctk.CTkButton(
            top_button_frame, text="フォルダから新規読み込み", command=self.scan_folder
            ).pack(side="left", padx=5)
        ctk.CTkButton(
            top_button_frame, text="CSVから読み込み", command=self.load_from_csv
            ).pack(side="left", padx=5)
        ctk.CTkButton(
            top_button_frame, text="フォルダと同期", command=self.sync_with_folder
            ).pack(side="left", padx=5)
        ctk.CTkButton(
            top_button_frame, text="CSVに保存", command=self.save_to_csv
            ).pack(side="left", padx=5)
        ctk.CTkButton(
            top_button_frame, text="統合PDFを生成", command=self.generate_pdf,
            fg_color="green", hover_color="darkgreen"
            ).pack(side="left", padx=10)

        self.scrollable_frame = ctk.CTkScrollableFrame(
            self, label_text="読み込み結果"
            )
        self.scrollable_frame.grid(
            row=2, column=0, padx=10, pady=10, sticky="nsew"
            )

    def get_icon_path(self):
        """
        実行環境(.exe or .py)に応じて、
        プロジェクトルートの 'assets' フォルダにある
        'synapsen.ico' のパスを返す。
        """
        try:
            if getattr(sys, 'frozen', False):
                # .exe実行の場合 (exeと同じフォルダがプロジェクトルート)
                project_root = Path(sys.executable).parent
            else:
                # .pyスクリプト実行の場合 (このファイルの親フォルダがプロジェクトルート)
                project_root = Path(__file__).parent.parent

            icon_path = project_root / 'assets' / 'synapsen.ico'

            if icon_path.is_file():
                return icon_path
        except Exception as e:
            print(f"Error finding icon path: {e}")
        return None

    def load_config(self):
        # 1. 実行ファイルの場所を基準としたbase_pathを最初に定義します
        if getattr(sys, 'frozen', False):
            # .exe実行の場合
            base_path = os.path.dirname(sys.executable)
        else:
            # .pyスクリプト実行の場合
            base_path = os.path.dirname(os.path.abspath(__file__))

        # 2. .exeか.pyかでconfig.iniの場所を決定します
        if getattr(sys, 'frozen', False):
            # .exe実行の場合（config.ini は .exe と同じフォルダ）
            config_path = os.path.join(base_path, 'config.ini')
        else:
            # スクリプト実行の場合（config.ini は .py の1つ上のフォルダ）
            config_path = os.path.join(
                os.path.abspath(os.path.join(base_path, '..')), 'config.ini'
                )
        print(f"[DEBUG] Loading config from: {config_path}")

        # config.ini があるフォルダのパスを基準として定義
        config_dir = os.path.dirname(config_path)

        config = configparser.ConfigParser(interpolation=None)

        # 3. configファイルが存在しない場合の処理
        if not os.path.exists(config_path):
            config['Paths'] = {
                'tags_data_path': 'tags.txt',
                'font_path': r'C:\windows\fonts\msgothic.ttc'
                }
            config['LaTeX'] = {
                'paper_size': "A4",
                'font': 'MS UI Gothic',
                'author': 'Your Name',
                'title_prefix': '月刊 統合ノート'
                }
            config['CommonplaceKeys'] = {
                'options': 'タスク,アイデア,思考・考察,コミュニケーション,学習・情報収集,日常・その他'
                }
            config['Extraction'] = {
                'key_rect': '26, 13, 400, 73'
                }
            config['KeyIcons'] = {
                'タスク': '♥',
                'アイデア': '♥',
                '思考・考察': '♥',
                'コミュニケーション': '♥',
                '学習・情報収集': '♥',
                '日常・その他': '♥'
                }
            config['KeyColors'] = {
                'タスク': 'FE0000',
                'アイデア': 'FFFF02',
                '思考・考察': '8802FF',
                'コミュニケーション': '02FF01',
                '学習・情報収集': '02FFFF',
                '日常・その他': 'F2F2F2'
                }
            with open(config_path, 'w', encoding='utf-8') as f:
                config.write(f)

        config.read(config_path, encoding='utf-8')

        # 4. configから読み込んだパスを、config.ini の場所 (config_dir) を基準に絶対パスへ変換します
        #    環境変数（%LOCALAPPDATA%など）も展開します
        font_path_from_config = config.get('Paths', 'font_path', fallback='')
        expanded_path = os.path.expandvars(font_path_from_config)  # 環境変数を展開

        if os.path.isabs(expanded_path):
            # configの値が絶対パス（または環境変数展開後、絶対パスになった）の場合、そのまま使用
            self.font_path = expanded_path
            # print(f"[DEBUG] Font path is ABSOLUTE: {self.font_path}")
        else:
            # configの値が相対パスの場合、config_dir と結合する
            self.font_path = os.path.join(config_dir, expanded_path)
            # print(f"[DEBUG] Font path is RELATIVE, resolved to: {self.font_path}")

        # 5. tags_data_pathの解決
        tags_path_from_config = config.get(
            'Paths', 'tags_data_path', fallback='tags.txt'
            )
        expanded_path = os.path.expandvars(tags_path_from_config)  # 環境変数を展開

        if os.path.isabs(expanded_path):
            self.tags_data_path = expanded_path
        else:
            # configの値が相対パスの場合、config_dir と結合する
            self.tags_data_path = os.path.join(
                config_dir, expanded_path
                )

        # 6. default_csv_path (追記先のマスターCSVパス) の解決
        default_csv_path_str = config.get(
            'Paths', 'default_csv_path', fallback=''
            )
        expanded_path = os.path.expandvars(default_csv_path_str)  # 環境変数を展開

        if not expanded_path:
            self.default_csv_path = None
            print("DEBUG: config.ini [Paths][default_csv_path] が未設定です。")
        elif os.path.isabs(expanded_path):
            self.default_csv_path = expanded_path
        else:
            # configの値が相対パスの場合、config_dir と結合する
            self.default_csv_path = os.path.join(config_dir, expanded_path)

        # 7. Automation設定の読み込み
        # 自動結合設定の読み込み
        self.auto_append_csv = config.getboolean(
            'Automation', 'auto_append_to_default_csv', fallback=False
            )
        if self.auto_append_csv and not self.default_csv_path:
            print("警告: auto_append_to_default_csv が True ですが、default_csv_path が未設定のため無効化されます。")
            self.auto_append_csv = False

        # 個別出力設定の読み込み
        self.create_individual_csv = config.getboolean(
            'Automation', 'create_individual_csv', fallback=False
            )

        self.paper_size = config.get(
            'LaTeX', 'paper_size', fallback='A4').upper()
        if self.paper_size == 'A5':
            self.paper_width = Helper.A5_WIDTH
            self.paper_height = Helper.A5_HEIGHT
            print("[DEBUG] Ersteller paper size set to A5")
        else:
            self.paper_size = 'A4'  # 不正な値はA4に
            self.paper_width = Helper.A4_WIDTH
            self.paper_height = Helper.A4_HEIGHT
            print("[DEBUG] Ersteller paper size set to A4")

        self.latex_font = config.get('LaTeX', 'font', fallback='Yu Gothic')
        self.latex_author = config.get('LaTeX', 'author', fallback='Your Name')

        self.latex_title_prefix = config.get(
            'LaTeX', 'title_prefix', fallback='月刊 統合ノート'
            )
        self.commonplace_key_options = [opt.strip() for opt in config.get('CommonplaceKeys', 'options', fallback='').split(',')]
        rect_str = config.get('Extraction', 'key_rect', fallback='0,0,0,0').split(',')
        self.key_rect = tuple(map(float, rect_str))
        self.key_icons = {k.lower(): v for k, v in config.items('KeyIcons')} if config.has_section('KeyIcons') else {}
        self.key_colors = {k.lower(): v for k, v in config.items('KeyColors')} if config.has_section('KeyColors') else {}

    def load_predefined_tags(self):
        try:
            tag_file = Path(self.tags_data_path)
            if tag_file.is_file():
                with open(tag_file, "r", encoding="utf-8") as f:
                    self.predefined_tags = [line.strip() for line in f if line.strip() and not line.startswith('#')]
                print(f"{len(self.predefined_tags)}件の事前定義タグを読み込みました。")
                # print(f"{self.tags_data_path}")
        except Exception as e:
            print(f"tags.txtの読み込み中にエラーが発生しました: {e}")

    def save_to_csv(self):
        if not self.all_notes_info:
            return
        filepath = tkinter.filedialog.asksaveasfilename(defaultextension=".csv", filetypes=[("CSV", "*.csv")], title="CSVファイルを保存")
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
                    "filepath"
                    ]
                writer = csv.DictWriter(
                    f, fieldnames=header, extrasaction='ignore'
                    )
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
                    row["tags"] = row.get("tags", "").split(";") if row.get("tags") else []
                    row["pages"] = int(row.get("pages", 0))
                    row["is_warning"] = row.get("date") in ["日付不明", "読み込み失敗"]
                    new_notes_info.append(row)
            self.all_notes_info = new_notes_info
            self.update_note_list()
            self.label.configure(text=f"読み込み完了: {os.path.basename(filepath)}")
        except Exception as e:
            self.label.configure(text=f"エラー: 読み込み失敗 - {e}")

    def append_to_master_csv(self, notes_to_append):
        """
        マスターCSV（config.iniのdefault_csv_path）に、
        ヘッダーを考慮しながらノート情報を追記する。
        """
        master_csv_path = Path(self.default_csv_path)
        
        # ファイルが存在し、中身が空でないかを確認
        file_exists_and_has_content = master_csv_path.is_file() and master_csv_path.stat().st_size > 0
        
        with open(master_csv_path, "a", newline="", encoding="utf-8-sig") as f:
            header = [
                "date", "time", "title", "pages", "tags",
                "key", "memo", "commonplace_key",
                "merged_pdf_filename", "merged_start_page"
            ]
            writer = csv.DictWriter(f, fieldnames=header, extrasaction='ignore')
            
            if not file_exists_and_has_content:
                writer.writeheader() # ファイルが新規 or 空の場合のみヘッダーを書き込む
                
            for note in notes_to_append:
                note_to_write = note.copy()
                note_to_write["tags"] = ";".join(sorted(note_to_write.get("tags", [])))
                writer.writerow(note_to_write)

    def save_merged_index_csv(self, notes_with_merged_info, merged_pdf_path):
        csv_filepath = Path(merged_pdf_path).with_suffix('.csv')
        try:
            with open(
                csv_filepath,
                "w",
                newline="",
                encoding="utf-8-sig"
            ) as f:

                header = [
                    "date",
                    "time",
                    "title",
                    "pages",
                    "tags",
                    "key",
                    "memo",
                    "commonplace_key",
                    "merged_pdf_filename",
                    "merged_start_page"
                ]
                writer = csv.DictWriter(
                    f,
                    fieldnames=header,
                    extrasaction='ignore'
                )
                writer.writeheader()
                for note in notes_with_merged_info:
                    note_to_write = note.copy()
                    note_to_write["tags"] =\
                        ";".join(sorted(note_to_write.get("tags", [])))
                    writer.writerow(note_to_write)
        except Exception as e:
            messagebox.showerror("CSV保存エラー", f"統合後目次CSVの保存に失敗しました: {e}")

    def scan_folder(self):
        folder_path = tkinter.filedialog.askdirectory(title="新規読み込みするフォルダを選択")
        if not folder_path:
            return
        self.label.configure(text=f"読み込み中: {folder_path}")
        self.update_idletasks()
        target_dir = Path(folder_path)
        self.all_notes_info = [
            info for pdf_file in target_dir.glob("*.pdf") if (info := Process.get_note_info(pdf_file, self.key_rect))
            ]

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
                parent_title = title[:-len(side_note_suffix)]

                # 親がマップに存在すれば、そのKeyを継承する
                if parent_title in parent_key_map:
                    info["commonplace_key"] = parent_key_map[parent_title]
                    keys_inherited_count += 1

        if keys_inherited_count > 0:
            print(f"DEBUG: {keys_inherited_count}件のサイドノートにIndex Keyを継承しました。")

        self.all_notes_info.sort(key=lambda note: (note['date'], note['time']))
        self.update_note_list()
        self.label.configure(text=f"読み込み完了！ {len(self.all_notes_info)}件のファイルを読み込みました。")

    def update_note_list(self):
        for widget in self.scrollable_frame.winfo_children():
            widget.destroy()
        if not self.all_notes_info:
            ctk.CTkLabel(
                self.scrollable_frame, text="PDFファイルが見つかりませんでした。"
            ).pack()
        else:
            default_text_color = ("#1F1F1F", "#1F1F1F")
            warning_text_color = ("#f08300", "#FF4500")
            for note_data in self.all_notes_info:
                row_frame = ctk.CTkFrame(
                    self.scrollable_frame,
                    fg_color="transparent"
                )
                cp_key = note_data.get('commonplace_key', '')
                icon = self.key_icons.get(cp_key.lower(), '')
                icon_color = self.key_colors.get(
                    cp_key.lower(),
                    default_text_color
                )
                if icon:
                    icon_label = ctk.CTkLabel(
                        row_frame,
                        text=icon, text_color=icon_color, font=("", 14)
                        )
                    icon_label.pack(side="left", padx=(0, 5))
                key_display = f" [ID: {note_data.get('key')}]" if note_data.get('key') else ""
                tags_display = " [タグ: " + ", ".join(sorted(note_data.get("tags", []))) + "]" if note_data.get("tags") else ""
                if note_data.get("is_warning"):
                    display_text = f"【警告】[{note_data.get('date')}] {note_data.get('title')}{key_display}{tags_display}"
                    text_color = warning_text_color
                else:
                    t = note_data.get('time', '')
                    time_display = f"({t[0:2]}:{t[2:4]}:{t[4:6]})" if t != "999999" else ""
                    display_text = f"日付: {note_data.get('date')} {time_display},{key_display} タイトル: {note_data.get('title')}{tags_display}"
                    text_color = default_text_color
                text_label = ctk.CTkLabel(
                    row_frame,
                    text=display_text, text_color=text_color, anchor="w"
                    )
                text_label.pack(side="left")
                command = lambda e, note=note_data: self.open_data_editor(note)
                row_frame.bind("<Button-1>", command)
                text_label.bind("<Button-1>", command)
                if 'icon_label' in locals() and icon_label.winfo_exists():
                    icon_label.bind("<Button-1>", command)
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
                if i + 1 < len(outline_items) and isinstance(outline_items[i+1], list):
                    # 子要素のリストに対して再帰的にこの関数を呼び出す
                    self._copy_bookmarks_recursive(
                        outline_items[i+1], writer, reader, parent=new_parent
                        )
                    i += 1  # 子要素リストをスキップするためインデックスを1つ進める
            i += 1

    def open_data_editor(self, note_data):
        session_tags = set()
        for note in self.all_notes_info:
            session_tags.update(note.get('tags', []))
        combined_tags = session_tags.union(set(self.predefined_tags))
        if hasattr(self, 'editor_window') and self.editor_window.winfo_exists():
            self.editor_window.focus()
            return
        self.editor_window = Dialogs.DataEditorWindow(
            self, note_data, list(combined_tags), self.commonplace_key_options
        )

    def generate_pdf(self):
        if not self.all_notes_info:
            self.label.configure(text="PDF生成対象のデータがありません。")
            return

        dialog = Dialogs.DateInputDialog(self)
        date_input = dialog.get_input()
        if not date_input:
            return
        year, month = date_input
        pdf_title = f"{self.latex_title_prefix} ({year}年 {month}月)"

        save_filepath = tkinter.filedialog.asksaveasfilename(
            defaultextension=".pdf",
            filetypes=[("PDFファイル", "*.pdf")],
            title="統合PDFの保存先を選択",
            initialfile=f"統合ノート_{year}_{month:02d}.pdf"
        )
        if not save_filepath:
            return

        self.label.configure(text="PDF生成中... しばらくお待ちください。")
        self.update_idletasks()

        temp_dir = tempfile.mkdtemp()
        try:
            # LaTeX生成に必要な設定情報を辞書にまとめる
            latex_config = {
                'latex_font': self.latex_font,
                'latex_author': self.latex_author,
                'key_icons': self.key_icons,
                'key_colors': self.key_colors
            }
            # 新しいモジュールの関数を呼び出す
            latex_source = Generator.create_latex_source(
                self.all_notes_info, latex_config, pdf_title, self.paper_size
            )

            tex_filepath = Path(temp_dir) / "mokuji.tex"
            with open(tex_filepath, "w", encoding="utf-8") as f:
                f.write(latex_source)

            self.label.configure(text="PDF生成中... (1/3) ページ構成を計算中")
            self.update_idletasks()

            for i in range(3):
                process = subprocess.run(
                    [
                        "lualatex",
                        "--shell-escape",
                        "-interaction=nonstopmode",
                        "mokuji.tex"
                    ],
                    cwd=temp_dir,
                    capture_output=True, text=True, encoding='utf-8',
                    errors='ignore'
                )
                if "Output written on" not in process.stdout:
                    print(f"--- LaTeX Compilation Error (Pass {i+1}) ---")
                    print(process.stdout)
                    print(process.stderr)
                    messagebox.showerror(
                        "LaTeX エラー",
                        f"PDFのコンパイルに失敗しました。(Pass {i+1})\n詳細はターミナルを確認してください。"
                    )
                    return

            draft_pdf_path = Path(temp_dir) / "mokuji.pdf"
            if not draft_pdf_path.is_file():
                messagebox.showerror("エラー", "LaTeXによる設計図PDFの生成に失敗しました。")
                return

            self.label.configure(text="PDF生成中... (2/3) ノートを結合中")
            self.update_idletasks()

            draft_reader = PdfReader(draft_pdf_path)
            final_writer = PdfWriter()

            note_content_start_page = -1

            first_note = self.all_notes_info[0]
            d = first_note['date']
            date_formatted = f"{d[0:4]}/{d[4:6]}/{d[6:8]}" if d.isdigit() and len(d) == 8 else d
            title_in_outline = first_note["title"].replace('_', ' ')

            expected_outline_title = f"{date_formatted} – {title_in_outline}"

            def find_title_in_outline(outline_items, target_title):
                for item in outline_items:
                    if isinstance(item, list):
                        result = find_title_in_outline(item, target_title)
                        if result:
                            return result
                    elif hasattr(item, 'title') and item.title.strip() == target_title.strip():
                        return item
                return None

            destination = find_title_in_outline(draft_reader.outline, expected_outline_title)

            # 本文の開始ページを取得
            if destination:
                note_content_start_page = draft_reader.get_destination_page_number(destination)
            else:
                messagebox.showerror(
                    "エラー",
                    f"設計図PDFの目次（しおり）から最初のノートの開始ページを見つけられませんでした。\n\n"
                    f"検索したタイトル:\n'{expected_outline_title}'\n\n"
                    "CSVやファイル名に特殊文字が含まれていないか確認してください。"
                )
                print("--- PDF Outline Search Debug ---")
                print(f"Searching for: '{expected_outline_title}'")
                print("Available outline items:")

                def print_outline(items, indent=0):
                    for item in items:
                        if isinstance(item, list):
                            print_outline(item, indent + 1)
                        elif hasattr(item, 'title'):
                            print('  ' * indent + f"- '{item.title}'")
                print_outline(draft_reader.outline)
                return

            index_start_page = -1
            # 「Index Key 索引」を検索して 索引の開始ページとして取得する
            for i in range(len(draft_reader.pages) - 1, -1, -1):
                page_text = draft_reader.pages[i].extract_text()
                if page_text and "Index Key 索引" in page_text:
                    index_start_page = i
                    break

            if index_start_page == -1:
                # フォールバックとして「タグ索引」も探す
                for i in range(len(draft_reader.pages) - 1, -1, -1):
                    page_text = draft_reader.pages[i].extract_text()
                    if page_text and "タグ索引" in page_text:
                        index_start_page = i
                        break

            if index_start_page == -1:
                messagebox.showerror("エラー", "設計図PDFから索引ページを特定できませんでした。")
                return

            note_total_pages = sum(note['pages'] for note in self.all_notes_info if Path(note.get("filepath", "")).is_file())
            if index_start_page - note_content_start_page != note_total_pages:
                messagebox.showwarning(
                    "ページ計算の警告",
                    f"計算されたページ数に矛盾があります。これは通常問題ありませんが、念のためご確認ください。\n\n"
                    f"本文の開始ページ: {note_content_start_page + 1}\n"
                    f"索引の開始ページ: {index_start_page + 1}\n"
                    f"確保されたページ数: {index_start_page - note_content_start_page}\n"
                    f"ノートの合計ページ数: {note_total_pages}\n\n"
                    "処理を続行します。"
                )

            for i in range(note_content_start_page):
                final_writer.add_page(draft_reader.pages[i])

            updated_notes_info = []
            note_page_cursor = note_content_start_page
            for note in self.all_notes_info:
                note['merged_start_page'] = note_page_cursor + 1
                note['merged_pdf_filename'] = Path(save_filepath).name
                updated_notes_info.append(note)

                if not Path(note.get("filepath", "")).is_file():
                    continue

                original_reader = PdfReader(note["filepath"])
                for i in range(len(original_reader.pages)):
                    if note_page_cursor >= index_start_page:
                        print(f"ページ数計算エラー:note_page_cursor({
                            note_page_cursor})が上限({index_start_page})を超えました。")
                        continue

                    template_page = draft_reader.pages[note_page_cursor]
                    content_page = original_reader.pages[i]

                    original_width = float(content_page.mediabox.width)
                    original_height = float(content_page.mediabox.height)
                    if original_width == 0 or original_height == 0:
                        continue

                    scale_w = self.paper_width / original_width
                    scale_h = self.paper_height / original_height
                    scale = min(scale_w, scale_h)
                    tx = (self.paper_width - original_width * scale) / 2
                    ty = (self.paper_height - original_height * scale) / 2
                    transform = Transformation().scale(
                        sx=scale, sy=scale
                        ).translate(
                            tx=tx, ty=ty
                            )
                    template_page.merge_transformed_page(
                        content_page, transform
                        )

                    final_writer.add_page(template_page)
                    note_page_cursor += 1

            for i in range(index_start_page, len(draft_reader.pages)):
                final_writer.add_page(draft_reader.pages[i])

            self.label.configure(text="PDF生成中... (3/3) 最終ファイル書き込み")
            self.update_idletasks()

            if draft_reader.outline:
                self._copy_bookmarks_recursive(
                    draft_reader.outline,
                    final_writer,
                    draft_reader
                    )

            with open(save_filepath, "wb") as f:
                final_writer.write(f)

            if self.auto_append_csv and self.default_csv_path:
                # --- A. 自動追記モード ---
                try:
                    self.append_to_master_csv(updated_notes_info)
                    
                    self.label.configure(text=f"成功！ 統合PDFを生成し、マスターCSVに追記しました。")
                    messagebox.showinfo(
                        "成功",
                        f"統合PDFの生成が完了しました。\n"
                        f"PDF: {os.path.basename(save_filepath)}\n\n"
                        f"目次情報は {os.path.basename(self.default_csv_path)} に自動追記されました。"
                    )
                except Exception as e:
                    messagebox.showerror("CSV追記エラー", f"マスターCSVへの追記に失敗しました: {self.default_csv_path}\n\n{e}")

            if self.create_individual_csv or not self.auto_append_csv:
                # --- B. 個別作成モード (自動追記が無効時 or 設定有効時) ---
                self.save_merged_index_csv(updated_notes_info, save_filepath)

                self.label.configure(text=f"成功！ 統合PDFと専用目次CSVを生成しました: {os.path.basename(save_filepath)}")
                messagebox.showinfo(
                    "成功",
                    "統合PDFと専用目次CSVの生成が完了しました。\n" +
                    f"PDF: {os.path.basename(save_filepath)}\n" +
                    f"CSV: {Path(save_filepath).with_suffix('.csv').name}"
                )

        finally:
            shutil.rmtree(temp_dir)

    def sync_with_folder(self):
        if not self.all_notes_info:
            self.label.configure(text="先にCSVを読み込んでください。")
            return
        folder_path = tkinter.filedialog.askdirectory(title="同期するフォルダを選択")
        if not folder_path:
            return
        self.label.configure(text=f"同期中: {folder_path}")
        self.update_idletasks()
        app_paths = {note.get('filepath') for note in self.all_notes_info}
        disk_paths = {
            str(pdf_file) for pdf_file in Path(folder_path).glob("*.pdf")
            }
        added_paths = disk_paths - app_paths
        deleted_paths = app_paths - disk_paths
        added_count, deleted_count = 0, 0
        if deleted_paths:
            deleted_filenames = "\n".join(
                [f"- {Path(p).name}" for p in deleted_paths]
                )
            user_response = messagebox.askyesno(
                "削除の確認",
                f"以下のファイルがフォルダから見つかりませんでした。リストから削除しますか？\n\n{deleted_filenames}"
                )
            if user_response:
                self.all_notes_info = [
                    note for note in self.all_notes_info if note.get(
                        'filepath'
                        ) not in deleted_paths
                    ]
                deleted_count = len(deleted_paths)
        if added_paths:
            for path in sorted(list(added_paths)):
                info = Process.get_note_info(Path(path), self.key_rect)
                if info:
                    self.all_notes_info.append(info)
            added_count = len(added_paths)
        if added_count > 0 or deleted_count > 0:
            self.all_notes_info.sort(
                key=lambda note: (note['date'], note['time'])
                )
            self.update_note_list()
            self.label.configure(
                text=f"同期完了！ {added_count}件追加, {deleted_count}件削除"
                )
        else:
            self.label.configure(text="変更はありませんでした。")


if __name__ == "__main__":
    app = Synapsen_Ersteller()

    if app.icon_path:  # <-- クラス内で取得したパスを利用
        try:
            # 'default=' を指定し、OSダイアログ(エクスプローラ等)にも適用
            app.iconbitmap(default=str(app.icon_path))
        except Exception as e:
            print(f"Icon default setting error: {e}")
    else:
        print("警告: アイコンファイル (assets/synapsen.ico) が見つかりません。")

    app.mainloop()

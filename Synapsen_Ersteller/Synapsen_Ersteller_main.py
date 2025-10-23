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
import datetime
import configparser

import PDFMargeHelper as Helper
import pdf_processor as Process
import latex_generator as Generator
import gui_dialogs as Dialogs


# ==============================================================================
# データ編集ウィンドウ
# ==============================================================================
class DataEditorWindow(ctk.CTkToplevel):
    def __init__(self, parent, note_data, all_tags, commonplace_key_options):
        super().__init__(parent)
        self.parent = parent
        self.note_data = note_data
        self.all_tags = all_tags
        self.temp_tags = list(self.note_data.get("tags", []))
        self.commonplace_key_options = commonplace_key_options

        self.title(f"データ編集: {self.note_data['title']}")
        self.geometry("500x700")
        self.transient(parent)
        self.grab_set()

        cp_key_frame = ctk.CTkFrame(self, fg_color="transparent")
        cp_key_frame.pack(pady=10, padx=10, fill="x")
        ctk.CTkLabel(
            cp_key_frame, text="Index Key 索引 (CP Key):", width=150, anchor="w"
            ).pack(side="left")
        self.cp_key_combo = ctk.CTkComboBox(
            cp_key_frame, values=self.commonplace_key_options
            )
        self.cp_key_combo.pack(side="left", expand=True, fill="x")
        self.cp_key_combo.set(self.note_data.get("commonplace_key", ""))

        key_frame = ctk.CTkFrame(self, fg_color="transparent")
        key_frame.pack(pady=10, padx=10, fill="x")
        ctk.CTkLabel(
            key_frame, text="ユニークID (Key):", width=150, anchor="w"
            ).pack(side="left")
        self.key_entry = ctk.CTkEntry(key_frame, placeholder_text="このノート固有のID")
        self.key_entry.pack(side="left", expand=True, fill="x")
        self.key_entry.insert(0, self.note_data.get("key", ""))

        memo_frame = ctk.CTkFrame(self, fg_color="transparent")
        memo_frame.pack(pady=10, padx=10, fill="both", expand=True)
        ctk.CTkLabel(memo_frame, text="要約・引用メモ:").pack(anchor="w")
        self.memo_textbox = ctk.CTkTextbox(memo_frame, height=150)
        self.memo_textbox.pack(fill="both", expand=True)
        self.memo_textbox.insert("1.0", self.note_data.get("memo", ""))

        tag_input_frame = ctk.CTkFrame(self, fg_color="transparent")
        tag_input_frame.pack(pady=5, padx=10, fill="x")
        ctk.CTkLabel(tag_input_frame, text="新しいタグ:").pack(side="left")
        self.tag_entry = ctk.CTkEntry(
            tag_input_frame, placeholder_text="Enterで追加"
            )
        self.tag_entry.pack(side="left", padx=5, expand=True, fill="x")
        self.tag_entry.bind("<Return>", self.add_tag_event)

        tag_button_frame = ctk.CTkFrame(self, fg_color="transparent")
        tag_button_frame.pack(pady=5, padx=10)
        ctk.CTkButton(
            tag_button_frame, text="タグを追加", command=self.add_tag_event
            ).pack(side="left", padx=5)
        ctk.CTkButton(
            tag_button_frame, text="既存タグから選択", command=self.open_tag_selector
            ).pack(side="left", padx=5)

        self.tags_frame = ctk.CTkScrollableFrame(self, label_text="現在のタグ")
        self.tags_frame.pack(pady=10, padx=10, fill="both", expand=True)

        bottom_button_frame = ctk.CTkFrame(self, fg_color="transparent")
        bottom_button_frame.pack(pady=10, side="bottom")
        ctk.CTkButton(
            bottom_button_frame, text="保存", command=self.save_and_close
            ).pack(side="left", padx=5)
        ctk.CTkButton(
            bottom_button_frame, text="キャンセル", command=self.destroy
            ).pack(side="left", padx=5)

        self.update_tags_display()

    def save_and_close(self):
        self.note_data["commonplace_key"] = self.cp_key_combo.get().strip()
        self.note_data["key"] = self.key_entry.get().strip()
        self.note_data["memo"] = self.memo_textbox.get("1.0", "end-1c").strip()
        self.note_data["tags"] = self.temp_tags
        self.parent.update_note_list()
        self.destroy()

    def update_tags_display(self):
        for widget in self.tags_frame.winfo_children():
            widget.destroy()
        for tag in sorted(self.temp_tags):
            tag_frame = ctk.CTkFrame(self.tags_frame)
            ctk.CTkLabel(tag_frame, text=tag).pack(side="left", padx=5)
            ctk.CTkButton(
                tag_frame, text="x", width=20,
                command=lambda t=tag: self.remove_tag(t)
                ).pack(side="left", padx=5)
            tag_frame.pack(anchor="w", pady=2, fill="x")

    def add_tag_event(self, event=None):
        new_tag = self.tag_entry.get().strip()
        if new_tag:
            parts = new_tag.split('_')
            for i in range(len(parts)):
                hierarchical_tag = "_".join(parts[:i+1])
                if hierarchical_tag not in self.temp_tags:
                    self.temp_tags.append(hierarchical_tag)
        self.update_tags_display()
        self.tag_entry.delete(0, "end")

    def remove_tag(self, tag_to_remove):
        self.temp_tags.remove(tag_to_remove)
        self.update_tags_display()

    def open_tag_selector(self):
        selector = TagSelectorWindow(self, self.all_tags, self.temp_tags)
        selected_tag = selector.get_selection()
        if selected_tag:
            self.tag_entry.delete(0, "end")
            self.tag_entry.insert(0, selected_tag)
            self.add_tag_event()


# ==============================================================================
# 既存タグ選択ウィンドウ
# ==============================================================================
class TagSelectorWindow(ctk.CTkToplevel):
    def __init__(self, parent, all_tags, current_tags):
        super().__init__(parent)
        self.selection = None
        self.title("既存のタグを選択")
        self.geometry("300x400")
        self.transient(parent)
        self.grab_set()

        scroll_frame = ctk.CTkScrollableFrame(self)
        scroll_frame.pack(fill="both", expand=True, padx=10, pady=10)
        tags_to_show = sorted(list(set(all_tags) - set(current_tags)))
        for tag in tags_to_show:
            btn = ctk.CTkButton(
                scroll_frame,
                text=tag, text_color=("#1F1F1F", "#1F1F1F"),
                fg_color="transparent", anchor="w",
                command=lambda t=tag: self.select_tag(t)
                )
            btn.pack(fill="x")

    def select_tag(self, tag):
        self.selection = tag
        self.destroy()

    def get_selection(self):
        self.master.wait_window(self)
        return self.selection


# ==============================================================================
# 年月入力ダイアログ
# ==============================================================================
class DateInputDialog(ctk.CTkToplevel):
    def __init__(self, parent):
        super().__init__(parent)
        self.title("年月を指定")
        self.geometry("300x200")
        self.result = None
        today = datetime.date.today()
        first_day_of_month = today.replace(day=1)
        last_month_date = first_day_of_month - datetime.timedelta(days=1)
        self.label = ctk.CTkLabel(self, text="生成するPDFの年月を入力してください:")
        self.label.pack(pady=10, padx=10)
        self.year_entry = ctk.CTkEntry(self, placeholder_text="年")
        self.year_entry.pack(pady=5)
        self.year_entry.insert(0, str(last_month_date.year))
        self.month_entry = ctk.CTkEntry(self, placeholder_text="月")
        self.month_entry.pack(pady=5)
        self.month_entry.insert(0, str(last_month_date.month))
        self.ok_button = ctk.CTkButton(self, text="OK", command=self.on_ok)
        self.ok_button.pack(pady=10)
        self.transient(parent)
        self.grab_set()

    def on_ok(self):
        try:
            year = int(self.year_entry.get())
            month = int(self.month_entry.get())
            if not (1 <= month <= 12):
                messagebox.showerror("入力エラー", "月は1から12の間で入力してください。")
                return
            self.result = (year, month)
            self.destroy()
        except ValueError:
            messagebox.showerror("入力エラー", "年と月には半角数字を入力してください。")

    def get_input(self):
        self.master.wait_window(self)
        return self.result


# ==============================================================================
# メインアプリケーションクラス
# ==============================================================================
class Synapsen_Ersteller(ctk.CTk):
    def __init__(self):
        super().__init__()
        self.title("Synapse Builder")
        self.geometry("800x700")
        self.grid_columnconfigure(0, weight=1)
        self.grid_rowconfigure(2, weight=1)

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

    def load_config(self):
        # 1. 実行ファイルの場所を基準としたbase_pathを最初に定義します
        if getattr(sys, 'frozen', False):
            base_path = os.path.dirname(sys.executable)
        else:
            base_path = os.path.dirname(os.path.abspath(__file__))

        # 2. base_pathを使ってconfig.iniの絶対パスを決定します
        config_path = os.path.join(os.path.abspath(os.path.join(base_path, '..')), 'config.ini')

        config = configparser.ConfigParser()

        # 3. configファイルが存在しない場合の処理
        if not os.path.exists(config_path):
            config['Paths'] = {
                'tags_data_path': 'tags.txt',
                'font_path': 'C:/Windows/Fonts/NotoSansJP-Regular.otf'
                }
            config['LaTeX'] = {
                'font': 'Yu Gothic',
                'author': 'Your Name',
                'title_prefix': '月刊 統合ノート'
                }
            config['CommonplaceKeys'] = {
                'options': '決意 / タスク・好奇心,勇気 / アイデア,正義 / 思考整理,親切 / コミュニケーション,誠実 / 学び・知識,不屈 / 考察・主観的記録,忍耐 / 記録・情報収集,無垢 / 日常・その他'
                }
            config['Extraction'] = {
                'key_rect': '26, 13, 400, 73'
                }
            config['KeyIcons'] = {
                '決意 / タスク・好奇心': '♥',
                '勇気 / アイデア': '♥',
                '正義 / 思考整理': '♥',
                '親切 / コミュニケーション': '♥',
                '誠実 / 学び・知識': '♥',
                '不屈 / 考察・主観的記録': '♥',
                '忍耐 / 記録・情報収集': '♥',
                '無垢 / 日常・その他': '♥'
                }
            config['KeyColors'] = {
                '決意 / タスク・好奇心': '#FE0000',
                '勇気 / アイデア': '#FF8000',
                '正義 / 思考整理': '#FFFF02',
                '親切 / コミュニケーション': '#02FF01',
                '誠実 / 学び・知識': '#0000FF',
                '不屈 / 考察・主観的記録': '#8802FF',
                '忍耐 / 記録・情報収集': '#02FFFF',
                '無垢 / 日常・その他': '#F2F2F2'
                }
            with open(config_path, 'w', encoding='utf-8') as f:
                config.write(f)

        config.read(config_path, encoding='utf-8')

        # 4. configから読み込んだ相対パスを、base_pathを基準に絶対パスへ変換します
        #    そして、他のメソッドで使えるように self.変数 に格納します
        font_path_from_config = config.get('Paths', 'font_path', fallback='')
        if os.path.isabs(font_path_from_config):
            # configの値が絶対パスの場合、そのまま使用する
            self.font_path = font_path_from_config
            # print(f"DEBUG: Font path is ABSOLUTE: {self.font_path}")
        else:
            # configの値が相対パスの場合、base_pathと結合する
            self.font_path = os.path.join(base_path, font_path_from_config)
            # print(f"DEBUG: Font path is RELATIVE, resolved to: {self.font_path}")

        # 5. tags_data_pathの解決
        tags_path_from_config = config.get(
            'Paths', 'tags_data_path', fallback='tags.txt'
            )
        if os.path.isabs(tags_path_from_config):
            self.tags_data_path = tags_path_from_config
        else:
            self.tags_data_path = os.path.join(
                base_path, tags_path_from_config
                )

        # 6. default_csv_path (追記先のマスターCSVパス) の解決
        default_csv_path_str = config.get('Paths', 'default_csv_path', fallback='')
        if not default_csv_path_str:
            self.default_csv_path = None
            print("DEBUG: config.ini [Paths][default_csv_path] が未設定です。")
        elif os.path.isabs(default_csv_path_str):
            self.default_csv_path = default_csv_path_str
        else:
            self.default_csv_path = os.path.join(base_path, default_csv_path_str)

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
                self.all_notes_info, latex_config, pdf_title
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
                        print(f"ページ数計算エラー: note_page_cursor({note_page_cursor})が上限({index_start_page})を超えました。")
                        continue

                    template_page = draft_reader.pages[note_page_cursor]
                    content_page = original_reader.pages[i]

                    original_width = float(content_page.mediabox.width)
                    original_height = float(content_page.mediabox.height)
                    if original_width == 0 or original_height == 0:
                        continue
                    scale_w = Helper.A4_WIDTH / original_width
                    scale_h = Helper.A4_HEIGHT / original_height
                    scale = min(scale_w, scale_h)
                    tx = (Helper.A4_WIDTH - original_width * scale) / 2
                    ty = (Helper.A4_HEIGHT - original_height * scale) / 2
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
    app.mainloop()

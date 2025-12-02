import os
import math
import datetime
import logging
from typing import List, Dict, Any

from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import A4, A5
from reportlab.lib.units import mm
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.lib.colors import HexColor

logger = logging.getLogger(__name__)

# --- カラー定義 ---
COLOR_KIKYO = HexColor("#5654a2")
COLOR_BLACK = HexColor("#000000")
COLOR_GRAY = HexColor("#888888")
COLOR_LIGHT_GRAY = HexColor("#cccccc")


class ReportLabPDFGenerator:
    """
    ReportLabを使用して、統合PDFの「骨格」（表紙、目次、ヘッダー/フッター、索引）
    を生成するクラス。
    """

    def __init__(self, config: Dict[str, Any]):
        """
        初期化処理。

        Args:
            config (Dict[str, Any]): アプリケーション設定辞書。
                                     フォントパス、アイコン設定、カラー設定などが含まれる。
        """
        self.config = config
        self.font_name = "SynapsenFont"
        self._register_font()

        # 設定からアイコンと色を取得（キーは小文字化して正規化）
        self.key_icons: Dict[str, str] = {
            k.lower(): v for k, v in config.get("key_icons", {}).items()
        }
        self.key_colors: Dict[str, str] = {
            k.lower(): v for k, v in config.get("key_colors", {}).items()
        }

    def _register_font(self) -> None:
        """
        PDF生成に使用するフォントを登録する。
        設定ファイル (config.ini) のパスを確認し、有効なフォントが見つからない場合は
        Helvetica (デフォルト) を使用する。
        """
        path_latex = self.config.get("latex_font", "")
        path_common = self.config.get("font_path", "")
        target_path = None

        if path_latex and os.path.exists(path_latex):
            target_path = path_latex
        elif path_common and os.path.exists(path_common):
            target_path = path_common
        else:
            logger.warning(
                "ReportLab: 有効なフォントパスが見つかりません。デフォルト(Helvetica)を使用します。"
            )
            self.font_name = "Helvetica"
            return

        try:
            # ReportLabは .ttf / .ttc を推奨。.otf は環境によってサポートが限定的
            if target_path.lower().endswith(".otf"):
                logger.warning(
                    f"ReportLab: .otf ファイルはサポート外の可能性があります: {target_path}"
                )
                pdfmetrics.registerFont(TTFont(self.font_name, target_path))
            else:
                pdfmetrics.registerFont(TTFont(self.font_name, target_path))
        except Exception as e:
            logger.error(f"ReportLab フォント登録エラー '{target_path}': {e}")
            self.font_name = "Helvetica"

    def _truncate_text(
        self,
        pdf_canvas: canvas.Canvas,
        text: str,
        max_width: float,
        font_name: str,
        font_size: float,
    ) -> str:
        """
        指定された幅に収まるようにテキストを切り詰め、末尾に "..." を付与する。

        Args:
            pdf_canvas (canvas.Canvas): 描画用キャンバスオブジェクト。
            text (str): 元のテキスト。
            max_width (float): 許容される最大幅 (ポイント単位)。
            font_name (str): フォント名。
            font_size (float): フォントサイズ。

        Returns:
            str: 切り詰められたテキスト。
        """
        if pdf_canvas.stringWidth(text, font_name, font_size) <= max_width:
            return text

        ellipsis = "..."
        max_width_without_ellipsis = max_width - pdf_canvas.stringWidth(
            ellipsis, font_name, font_size
        )

        # 文字列を後ろから削って幅をチェック
        current_length = len(text)
        while current_length > 0:
            sub_text = text[:current_length]
            if (
                pdf_canvas.stringWidth(sub_text, font_name, font_size)
                <= max_width_without_ellipsis
            ):
                return sub_text + ellipsis
            current_length -= 1

        return ellipsis

    def create_skeleton_pdf(
        self,
        notes_info: List[Dict[str, Any]],
        title: str,
        paper_size: str,
        output_path: str,
    ) -> Dict[str, int]:
        """
        目次、ヘッダー、フッター、索引を含む「骨格PDF」を生成する。
        実際のノート本文（中身）はここには描画されず、後工程で結合される。

        Args:
            notes_info (List[Dict[str, Any]]): 各ノートのメタデータリスト。
            title (str): PDFのタイトル。
            paper_size (str): "A4" または "A5"。
            output_path (str): 生成するPDFの保存パス。

        Returns:
            Dict[str, int]: コンテンツ開始ページ番号などのメタデータ辞書。
                            {'content_start_page': int, 'index_start_page': int}
        """
        pagesize = A5 if paper_size.upper() == "A5" else A4
        page_width, page_height = pagesize

        # 目次の行の高さ概算
        line_height = 6 * mm
        lines_per_page = int((page_height - 50 * mm) / line_height)

        # 1. データ準備 (ページ数0のノートは除外)
        valid_notes = [n for n in notes_info if n.get("pages", 0) > 0]

        # 索引データの構築
        commonplace_key_map = self._build_key_map(valid_notes)
        tag_map = self._build_tag_map(valid_notes)

        has_cp_index = bool(commonplace_key_map)
        has_tag_index = bool(tag_map)

        # 2. ページ数の計算と予測

        # 目次ページ数の計算
        toc_extra_lines = 0
        if has_cp_index:
            toc_extra_lines += 1
        if has_tag_index:
            toc_extra_lines += 1

        toc_total_lines = len(valid_notes) + toc_extra_lines
        toc_pages_count = (
            math.ceil(toc_total_lines / lines_per_page) if toc_total_lines > 0 else 1
        )

        # 本文の開始ページ (表紙(1) + 目次ページ数 + 1)
        content_start_page_num = 1 + toc_pages_count + 1

        # 各ノートの開始ページ番号をマップ (Index: PageNum)
        note_page_map = {}
        cursor_page = content_start_page_num
        for i, note in enumerate(valid_notes):
            note_page_map[i] = cursor_page
            cursor_page += note.get("pages", 0)

        index_start_page_num = cursor_page

        # 索引ページ数の見積もり
        cp_key_index_pages_count = 0
        if has_cp_index:
            cp_key_index_pages_count = self._calculate_index_pages(
                commonplace_key_map, page_height
            )

        tag_index_start_page_num = index_start_page_num + cp_key_index_pages_count

        # 3. 描画開始
        pdf_canvas = canvas.Canvas(str(output_path), pagesize=pagesize)
        pdf_canvas.setTitle(title)

        # (A) 表紙の描画
        self._draw_cover(pdf_canvas, title, page_width, page_height)
        pdf_canvas.showPage()

        # (B) 目次の描画
        pdf_canvas.bookmarkPage("DEST_TOC")
        pdf_canvas.addOutlineEntry("目次", "DEST_TOC")

        self._draw_toc(
            pdf_canvas,
            valid_notes,
            note_page_map,
            page_width,
            page_height,
            toc_pages_count,
            index_key_start_page=index_start_page_num,
            tag_index_start_page=tag_index_start_page_num,
            has_cp_index=has_cp_index,
            has_tag_index=has_tag_index,
        )

        # 目次後の空白ページ調整 (本文開始ページまで空ページを送る)
        while pdf_canvas.getPageNumber() < content_start_page_num:
            pdf_canvas.showPage()

        # (C) 本文エリア (ヘッダー・フッターのみ描画)
        for i, note in enumerate(valid_notes):
            page_count = note.get("pages", 0)

            # ノートへの内部リンク先(Destination)を作成
            dest_key = f"NOTE_LINK_{i}"
            pdf_canvas.bookmarkPage(dest_key)

            # PDFのしおり(Outline)に追加
            title_text = note.get("title", "No Title")
            date_text = note.get("date", "")
            if len(date_text) == 8:
                date_text = f"{date_text[:4]}/{date_text[4:6]}/{date_text[6:]}"
            bookmark_title = f"{date_text} {title_text}"
            pdf_canvas.addOutlineEntry(bookmark_title, dest_key)

            # ノートの各ページに対してヘッダー/フッターを描画
            for p_idx in range(1, page_count + 1):
                self._draw_header(
                    pdf_canvas, note, p_idx, page_count, page_width, page_height
                )
                # フッターには物理ページ番号(通し番号)を使用
                self._draw_footer(
                    pdf_canvas, pdf_canvas.getPageNumber(), page_width, page_height
                )
                pdf_canvas.showPage()

        # (D) Index Key 索引の描画
        if has_cp_index:
            pdf_canvas.bookmarkPage("DEST_INDEX_KEY")
            pdf_canvas.addOutlineEntry("Index Key 索引", "DEST_INDEX_KEY")

            self._draw_index_page(
                pdf_canvas,
                "Index Key 索引",
                sorted(commonplace_key_map.keys()),
                commonplace_key_map,
                valid_notes,
                note_page_map,
                page_width,
                page_height,
                use_color=True,
                group_by_hierarchy=True,
            )

        # (E) タグ索引の描画
        if has_tag_index:
            pdf_canvas.bookmarkPage("DEST_INDEX_TAG")
            pdf_canvas.addOutlineEntry("タグ索引", "DEST_INDEX_TAG")

            self._draw_index_page(
                pdf_canvas,
                "タグ索引",
                sorted(tag_map.keys()),
                tag_map,
                valid_notes,
                note_page_map,
                page_width,
                page_height,
                use_color=False,
                group_by_hierarchy=True,
            )

        pdf_canvas.save()
        logger.info("ReportLab: Skeleton PDF saved successfully.")

        return {
            "content_start_page": content_start_page_num - 1,
            "index_start_page": index_start_page_num - 1,
        }

    # --- ヘルパーメソッド ---

    def _build_key_map(self, notes: List[Dict[str, Any]]) -> Dict[str, List[int]]:
        """ノートリストからIndex Keyごとのインデックスリストを作成する"""
        key_map = {}
        for i, note in enumerate(notes):
            cp_key = note.get("commonplace_key", "")
            if not cp_key:
                cp_key = "（未分類）"
            if cp_key not in key_map:
                key_map[cp_key] = []
            key_map[cp_key].append(i)
        return key_map

    def _build_tag_map(self, notes: List[Dict[str, Any]]) -> Dict[str, List[int]]:
        """ノートリストからタグごとのインデックスリストを作成する"""
        tag_map = {}
        for i, note in enumerate(notes):
            tags = note.get("tags", [])
            # 文字列の場合はリストに変換
            if isinstance(tags, str):
                tags = [t.strip() for t in tags.split(";") if t.strip()]

            for tag in tags:
                if not tag:
                    continue
                if tag not in tag_map:
                    tag_map[tag] = []
                tag_map[tag].append(i)
        return tag_map

    def _calculate_index_pages(
        self, data_map: Dict[str, List[int]], page_height: float
    ) -> int:
        """
        索引に必要なページ数を概算する。

        Args:
            data_map: {キー: [ノートインデックス, ...]} 形式の辞書
            page_height: ページ高さ
        """
        if not data_map:
            return 0

        margin_y_top = 30 * mm
        margin_y_bottom = 20 * mm
        line_height = 5 * mm
        usable_height = page_height - margin_y_top - margin_y_bottom

        current_y = usable_height - 10 * mm
        page_count = 1
        sorted_keys = sorted(data_map.keys())

        for key in sorted_keys:
            # 見出しに必要なスペースがない場合は改ページ
            if current_y < line_height * 2:
                page_count += 1
                current_y = usable_height

            current_y -= line_height  # 見出し行

            for _ in data_map[key]:
                # 各行の描画スペース確認
                if current_y < line_height:
                    page_count += 1
                    current_y = usable_height
                current_y -= line_height

            current_y -= 2 * mm  # グループ間のスペース

        return page_count

    # --- 描画メソッド ---

    def _draw_cover(
        self, pdf_canvas: canvas.Canvas, title: str, width: float, height: float
    ) -> None:
        """表紙を描画する"""
        pdf_canvas.setFont(self.font_name, 24)
        pdf_canvas.setFillColor(COLOR_BLACK)
        pdf_canvas.drawCentredString(width / 2, height / 2 + 20 * mm, title)

        pdf_canvas.setFont(self.font_name, 12)
        author = self.config.get("latex_author", "Synapsen Ersteller")
        pdf_canvas.drawCentredString(width / 2, height / 2 - 20 * mm, author)

        dt_now = datetime.datetime.now()
        date_str = dt_now.strftime("%Y年%m月%d日 生成")
        pdf_canvas.setFont(self.font_name, 10)
        pdf_canvas.setFillColor(COLOR_GRAY)
        pdf_canvas.drawCentredString(width / 2, height / 2 - 30 * mm, date_str)

    def _draw_toc(
        self,
        pdf_canvas: canvas.Canvas,
        notes: List[Dict[str, Any]],
        page_map: Dict[int, int],
        width: float,
        height: float,
        expected_pages: int,
        index_key_start_page: int,
        tag_index_start_page: int,
        has_cp_index: bool,
        has_tag_index: bool,
    ) -> None:
        """
        目次ページを描画する。
        """
        margin_x = 20 * mm
        margin_y_top = 30 * mm
        line_height = 6 * mm
        current_y = height - margin_y_top

        # タイトル
        pdf_canvas.setFillColor(COLOR_BLACK)
        pdf_canvas.setFont(self.font_name, 16)
        pdf_canvas.drawString(margin_x, current_y, "目次")
        current_y -= 15 * mm
        pdf_canvas.setFont(self.font_name, 10)

        start_page_of_toc = pdf_canvas.getPageNumber()
        page_num_width = 15 * mm
        max_text_width = width - (margin_x * 2) - page_num_width - 5 * mm

        # ノートリストの描画
        for i, note in enumerate(notes):
            # 改ページ判定
            if current_y < 20 * mm:
                pdf_canvas.showPage()
                current_y = height - margin_y_top
                pdf_canvas.setFont(self.font_name, 10)

            title = note.get("title", "")
            date = note.get("date", "")
            if len(date) == 8:
                date = f"{date[:4]}/{date[4:6]}/{date[6:]}"

            raw_text = f"{date}  {title}"
            text = self._truncate_text(
                pdf_canvas, raw_text, max_text_width, self.font_name, 10
            )

            # タイトル描画
            pdf_canvas.setFillColor(COLOR_KIKYO)
            pdf_canvas.drawString(margin_x, current_y, text)

            # ページ番号描画
            pdf_canvas.setFillColor(COLOR_BLACK)
            dest_page = page_map[i]
            page_str = str(dest_page)
            pdf_canvas.drawRightString(width - margin_x, current_y, page_str)

            # リーダー線 (点線)
            pdf_canvas.setFillColor(COLOR_BLACK)
            text_w = pdf_canvas.stringWidth(text, self.font_name, 10) + 1 * mm
            pdf_canvas.saveState()
            pdf_canvas.setDash(1, 2)
            pdf_canvas.setLineWidth(0.5)
            pdf_canvas.line(
                margin_x + text_w + 2 * mm,
                current_y + 1 * mm,
                width - margin_x - 10 * mm,
                current_y + 1 * mm,
            )
            pdf_canvas.restoreState()

            # 内部リンク領域の設定
            link_rect = (margin_x, current_y - 2, width - margin_x, current_y + 10)
            pdf_canvas.linkRect("", destinationname=f"NOTE_LINK_{i}", Rect=link_rect)

            current_y -= line_height

        # --- Index Key 索引へのリンク ---
        if has_cp_index:
            if current_y < 20 * mm:
                pdf_canvas.showPage()
                current_y = height - margin_y_top
                pdf_canvas.setFont(self.font_name, 10)

            pdf_canvas.setFillColor(COLOR_KIKYO)
            pdf_canvas.drawString(margin_x, current_y, "Index Key 索引")

            pdf_canvas.setFillColor(COLOR_BLACK)
            pdf_canvas.drawRightString(
                width - margin_x, current_y, str(index_key_start_page)
            )

            text_w = (
                pdf_canvas.stringWidth("Index Key 索引", self.font_name, 10) + 1 * mm
            )
            pdf_canvas.saveState()
            pdf_canvas.setDash(1, 2)
            pdf_canvas.setLineWidth(0.5)
            pdf_canvas.line(
                margin_x + text_w + 2 * mm,
                current_y + 1 * mm,
                width - margin_x - 10 * mm,
                current_y + 1 * mm,
            )
            pdf_canvas.restoreState()

            link_rect = (margin_x, current_y - 2, width - margin_x, current_y + 10)
            pdf_canvas.linkRect("", destinationname="DEST_INDEX_KEY", Rect=link_rect)
            current_y -= line_height

        # --- タグ索引へのリンク ---
        if has_tag_index:
            if current_y < 20 * mm:
                pdf_canvas.showPage()
                current_y = height - margin_y_top
                pdf_canvas.setFont(self.font_name, 10)

            pdf_canvas.setFillColor(COLOR_KIKYO)
            pdf_canvas.drawString(margin_x, current_y, "タグ索引")

            pdf_canvas.setFillColor(COLOR_BLACK)
            pdf_canvas.drawRightString(
                width - margin_x, current_y, str(tag_index_start_page)
            )

            text_w = pdf_canvas.stringWidth("タグ索引", self.font_name, 10) + 1 * mm
            pdf_canvas.saveState()
            pdf_canvas.setDash(1, 2)
            pdf_canvas.setLineWidth(0.5)
            pdf_canvas.line(
                margin_x + text_w + 2 * mm,
                current_y + 1 * mm,
                width - margin_x - 10 * mm,
                current_y + 1 * mm,
            )
            pdf_canvas.restoreState()

            link_rect = (margin_x, current_y - 2, width - margin_x, current_y + 10)
            pdf_canvas.linkRect("", destinationname="DEST_INDEX_TAG", Rect=link_rect)
            current_y -= line_height

        # 予定ページ数に達するまで空ページを追加 (本文開始位置の調整)
        while (pdf_canvas.getPageNumber() - start_page_of_toc + 1) < expected_pages:
            pdf_canvas.showPage()

        # 目次の終了
        pdf_canvas.showPage()

    def _draw_header(
        self,
        pdf_canvas: canvas.Canvas,
        note: Dict[str, Any],
        current_page: int,
        total_pages: int,
        width: float,
        height: float,
    ) -> None:
        """各ページの上部にヘッダー（日付、タイトル、アイコン）を描画する"""
        margin_x = 20 * mm
        header_y = height - 15 * mm

        title = note.get("title", "")
        date = note.get("date", "")
        if len(date) == 8:
            date = f"{date[:4]}/{date[4:6]}/{date[6:]}"

        cp_key = note.get("commonplace_key", "").lower()
        icon = self.key_icons.get(cp_key, "")
        color_hex = self.key_colors.get(cp_key, "#000000")

        pdf_canvas.saveState()
        pdf_canvas.setLineWidth(0.5)
        pdf_canvas.setStrokeColor(COLOR_GRAY)
        pdf_canvas.line(
            margin_x, header_y - 2 * mm, width - margin_x, header_y - 2 * mm
        )

        current_x = margin_x
        # アイコン描画
        if icon:
            try:
                pdf_canvas.setFillColor(HexColor(color_hex))
                pdf_canvas.setFont(self.font_name, 12)
                pdf_canvas.drawString(current_x, header_y, icon)
                current_x += pdf_canvas.stringWidth(icon, self.font_name, 12) + 2 * mm
            except Exception:
                pass

        # テキスト描画
        pdf_canvas.setFillColor(COLOR_BLACK)
        pdf_canvas.setFont(self.font_name, 9)
        raw_text = f"{date}  {title}"
        max_header_w = width - current_x - margin_x - 30 * mm
        text = self._truncate_text(
            pdf_canvas, raw_text, max_header_w, self.font_name, 9
        )

        pdf_canvas.drawString(current_x, header_y, text)

        # ノート内ページ番号 (例: 1/3)
        pdf_canvas.drawRightString(
            width - margin_x, header_y, f"({current_page}/{total_pages})"
        )
        pdf_canvas.restoreState()

    def _draw_footer(
        self, pdf_canvas: canvas.Canvas, page_num: int, width: float, height: float
    ) -> None:
        """各ページの下部にフッター（物理ページ番号）を描画する"""
        footer_y = 10 * mm
        pdf_canvas.saveState()
        margin_x = 20 * mm
        line_y = footer_y + 5 * mm

        pdf_canvas.setLineWidth(0.5)
        pdf_canvas.setStrokeColor(COLOR_GRAY)
        pdf_canvas.line(margin_x, line_y, width - margin_x, line_y)

        pdf_canvas.setFillColor(COLOR_BLACK)
        pdf_canvas.setFont(self.font_name, 10)
        pdf_canvas.drawCentredString(width / 2, footer_y, str(page_num))
        pdf_canvas.restoreState()

    def _draw_index_page(
        self,
        pdf_canvas: canvas.Canvas,
        title: str,
        sorted_keys: List[str],
        data_map: Dict[str, List[int]],
        notes: List[Dict[str, Any]],
        page_map: Dict[int, int],
        width: float,
        height: float,
        use_color: bool = False,
        group_by_hierarchy: bool = False,
    ) -> None:
        """
        索引ページ（Index Key または タグ）を描画する。
        """
        margin_x = 20 * mm
        margin_y_top = 30 * mm
        current_y = height - margin_y_top
        line_height = 5 * mm

        # タイトル
        pdf_canvas.setFillColor(COLOR_BLACK)
        pdf_canvas.setFont(self.font_name, 16)
        pdf_canvas.drawString(margin_x, current_y, title)
        current_y -= 10 * mm

        prev_top_level = None
        list_indent_x = margin_x + 9 * mm
        max_list_text_w = width - list_indent_x - margin_x - 15 * mm

        for key in sorted_keys:
            # 階層表示のグループ分け (例: "Work/ProjectA" -> "Work" で空行を入れる)
            if group_by_hierarchy:
                current_top = key.split("/")[0] if "/" in key else key
                if prev_top_level is not None and current_top != prev_top_level:
                    current_y -= line_height
                prev_top_level = current_top

            # 改ページ判定
            if current_y < 20 * mm:
                pdf_canvas.showPage()
                current_y = height - margin_y_top

            # キーの見出し描画
            pdf_canvas.setFont(self.font_name, 11)
            pdf_canvas.setFillColor(COLOR_BLACK)

            display_text = f"{key}"

            # 色とアイコンの適用 (Index Keyの場合)
            if use_color:
                key_lower = key.lower()
                color_hex = self.key_colors.get(key_lower, "#000000")
                icon = self.key_icons.get(key_lower, "")
                try:
                    pdf_canvas.setFillColor(HexColor(color_hex))
                except Exception:
                    pass
                if icon:
                    display_text = f"{icon} {key}"

            pdf_canvas.drawString(margin_x, current_y, display_text)
            pdf_canvas.setFillColor(COLOR_BLACK)
            current_y -= line_height + 2 * mm

            # キーに紐づくノートリストの描画
            pdf_canvas.setFont(self.font_name, 9)
            note_indices = data_map[key]

            for idx in note_indices:
                if current_y < 20 * mm:
                    pdf_canvas.showPage()
                    current_y = height - margin_y_top
                    pdf_canvas.setFont(self.font_name, 9)

                note = notes[idx]
                note_title = note.get("title", "")
                dest_page = page_map[idx]

                raw_title_text = f"{note_title}, "
                line_text_left = self._truncate_text(
                    pdf_canvas, raw_title_text, max_list_text_w, self.font_name, 9
                )

                pdf_canvas.setFillColor(COLOR_BLACK)
                pdf_canvas.drawString(list_indent_x, current_y, line_text_left)

                page_str = f"p.{dest_page}"
                text_width = pdf_canvas.stringWidth(line_text_left, self.font_name, 9)
                page_x = list_indent_x + text_width + 2

                pdf_canvas.setFillColor(COLOR_KIKYO)
                pdf_canvas.drawString(page_x, current_y, page_str)

                # リンク設定
                total_width = (
                    text_width + pdf_canvas.stringWidth(page_str, self.font_name, 9) + 5
                )
                link_rect = (
                    margin_x,
                    current_y - 1,
                    list_indent_x + total_width,
                    current_y + 8,
                )

                pdf_canvas.linkRect(
                    "", destinationname=f"NOTE_LINK_{idx}", Rect=link_rect
                )

                current_y -= line_height

            current_y -= 2 * mm

        pdf_canvas.showPage()

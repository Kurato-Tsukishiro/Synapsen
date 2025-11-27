import customtkinter as ctk
import tkinter as tk
from PIL import Image, ImageTk
import numpy as np
from pathlib import Path


class PerspectiveCropEditor(ctk.CTkToplevel):
    def __init__(self, parent, image_data, on_save_callback):
        """
        4隅指定による台形補正エディタ (ズーム・パン・辺移動[軸固定]対応版)
        """
        super().__init__(parent)

        # --- アイコン設定 ---
        self._custom_icon_path = None
        # 親ウィンドウ等からアイコンパスの取得を試みる
        if hasattr(parent, "_custom_icon_path") and parent._custom_icon_path:
            self._custom_icon_path = parent._custom_icon_path
        elif hasattr(parent, "icon_path") and parent.icon_path:
            self._custom_icon_path = str(parent.icon_path)
        elif (
            hasattr(parent, "parent_app")
            and hasattr(parent.parent_app, "icon_path")
            and parent.parent_app.icon_path
        ):
            self._custom_icon_path = str(parent.parent_app.icon_path)

        if self._custom_icon_path:
            # ウィンドウ生成直後のリセットを防ぐため、少し遅延させて適用
            self.after(200, lambda: self.iconbitmap(default=self._custom_icon_path))
        # -------------------------

        self.title("画像変形・トリミング")
        self.geometry("1000x800")
        self.on_save_callback = on_save_callback

        # 画像の読み込み
        self.original_image = self._load_image(image_data)
        self.display_image = None
        self.tk_image = None

        # 表示制御パラメータ
        self.base_scale = 1.0  # ウィンドウに合わせるための基本倍率
        self.zoom_level = 1.0  # ユーザーによるズーム倍率
        self.pan_x = 0  # パンによるオフセットX
        self.pan_y = 0  # パンによるオフセットY
        self.view_offset_x = 0  # 中央寄せのための基本オフセットX
        self.view_offset_y = 0  # 中央寄せのための基本オフセットY

        # 4隅の座標 (画像座標系: 左上, 右上, 右下, 左下)
        self.corners = []

        # ドラッグ操作用
        self.drag_state = {
            "item": None,  # 操作中のハンドルID (0-3:角, 4-7:辺)
            "type": None,  # "corner" or "edge" or "pan"
            "last_x": 0,  # 前回のマウスX (パン/移動用)
            "last_y": 0,  # 前回のマウスY
        }

        # デザイン設定
        self.corner_radius = 8
        self.edge_radius = 6
        self.magnifier_size = 150
        self.magnifier_zoom = 2.5
        self.magnifier_image_ref = None

        # UI構築
        self._create_widgets()

        # 初期化遅延実行 (画面描画後にキャンバスサイズを取得するため)
        self.after(100, self._init_view)

        self.grab_set()

    def _load_image(self, data):
        if isinstance(data, Path):
            return Image.open(data).convert("RGB")
        elif isinstance(data, Image.Image):
            return data.convert("RGB")
        else:
            raise ValueError("Unsupported image data type")

    def _create_widgets(self):
        # ツールバー
        self.toolbar = ctk.CTkFrame(self)
        self.toolbar.pack(side="bottom", fill="x", padx=10, pady=10)

        # 操作説明ラベル
        ctk.CTkLabel(
            self.toolbar,
            text="ホイール: ズーム | 右ドラッグ: 移動 | ■ハンドル: 辺移動(軸固定)",
            font=("", 12),
            text_color="gray",
        ).pack(side="left", padx=15)

        self.btn_cancel = ctk.CTkButton(
            self.toolbar,
            text="キャンセル",
            fg_color="gray",
            width=80,
            command=self.destroy,
        )
        self.btn_cancel.pack(side="left", padx=5)

        self.btn_reset = ctk.CTkButton(
            self.toolbar,
            text="リセット (全体表示)",
            width=120,
            command=self._reset_view_and_corners,
        )
        self.btn_reset.pack(side="left", padx=5)

        self.btn_save = ctk.CTkButton(
            self.toolbar, text="変形して適用", width=120, command=self._apply_and_save
        )
        self.btn_save.pack(side="right", padx=5)

        # キャンバス
        self.canvas_frame = ctk.CTkFrame(self)
        self.canvas_frame.pack(side="top", fill="both", expand=True, padx=10, pady=10)

        self.canvas = tk.Canvas(
            self.canvas_frame, bg="#333333", highlightthickness=0, cursor="crosshair"
        )
        self.canvas.pack(fill="both", expand=True)

        # イベントバインド
        # 左クリック: ハンドル操作
        self.canvas.bind("<ButtonPress-1>", self._on_left_down)
        self.canvas.bind("<ButtonRelease-1>", self._on_left_up)
        self.canvas.bind("<B1-Motion>", self._on_left_drag)

        # 右クリック / ホイールクリック: パン操作
        self.canvas.bind("<ButtonPress-3>", self._on_pan_start)
        self.canvas.bind("<B3-Motion>", self._on_pan_drag)
        # マウスホイール押し込み対応
        self.canvas.bind("<ButtonPress-2>", self._on_pan_start)
        self.canvas.bind("<B2-Motion>", self._on_pan_drag)

        # ホイール: ズーム
        self.canvas.bind("<MouseWheel>", self._on_wheel)  # Windows
        self.canvas.bind("<Button-4>", self._on_wheel)  # Linux up
        self.canvas.bind("<Button-5>", self._on_wheel)  # Linux down

        # リサイズ
        self.bind("<Configure>", self._on_resize)

    # --- 座標変換ヘルパー ---
    @property
    def current_scale(self):
        return self.base_scale * self.zoom_level

    def _img_to_canvas(self, ix, iy):
        """画像座標(px) -> キャンバス座標(px)"""
        cx = (ix * self.current_scale) + self.view_offset_x + self.pan_x
        cy = (iy * self.current_scale) + self.view_offset_y + self.pan_y
        return cx, cy

    def _canvas_to_img(self, cx, cy):
        """キャンバス座標(px) -> 画像座標(px)"""
        ix = (cx - self.view_offset_x - self.pan_x) / self.current_scale
        iy = (cy - self.view_offset_y - self.pan_y) / self.current_scale
        return ix, iy

    # --- ビュー初期化・更新 ---
    def _init_view(self):
        """画像をキャンバスにフィットさせる初期パラメータを計算"""
        self.update_idletasks()
        cw = self.canvas.winfo_width()
        ch = self.canvas.winfo_height()
        if cw <= 1:
            return

        iw, ih = self.original_image.size

        # フィットする倍率
        scale = min(cw / iw, ch / ih) * 0.9
        self.base_scale = scale
        self.zoom_level = 1.0
        self.pan_x = 0
        self.pan_y = 0

        # 中央寄せオフセット
        self.view_offset_x = (cw - iw * scale) / 2
        self.view_offset_y = (ch - ih * scale) / 2

        # 初回のみ四隅を初期化
        if not self.corners:
            self.corners = [[0, 0], [iw, 0], [iw, ih], [0, ih]]

        self._redraw_all()

    def _reset_view_and_corners(self):
        self.corners = []
        self._init_view()

    def _on_resize(self, event):
        if event.widget == self:
            self._redraw_all()

    def _redraw_all(self):
        """画像とオーバーレイの再描画"""
        self.canvas.delete("all")

        iw, ih = self.original_image.size

        # 1. 画像の描画
        display_w = int(iw * self.current_scale)
        display_h = int(ih * self.current_scale)

        if display_w > 0 and display_h > 0:
            if self.display_image is None or self.display_image.size != (
                display_w,
                display_h,
            ):
                self.display_image = self.original_image.resize(
                    (display_w, display_h), Image.Resampling.BILINEAR
                )
                self.tk_image = ImageTk.PhotoImage(self.display_image)

            # 表示位置
            img_cx, img_cy = self._img_to_canvas(0, 0)
            self.canvas.create_image(
                img_cx, img_cy, anchor="nw", image=self.tk_image, tags="image"
            )

        # 2. オーバーレイ（枠・ハンドル）の描画
        self._draw_overlays()

    def _draw_overlays(self):
        # 初期化前(cornersが空)の場合は描画しない (IndexError回避)
        if not self.corners:
            return

        self.canvas.delete("ui")

        # コーナーのキャンバス座標リストを作成
        flat_c_pts = []
        c_pts_tuples = []

        for x, y in self.corners:
            cx, cy = self._img_to_canvas(x, y)
            flat_c_pts.extend([cx, cy])
            c_pts_tuples.append((cx, cy))

        # 枠線
        self.canvas.create_polygon(
            flat_c_pts, outline="#00E5FF", width=2, fill="", tags="ui"
        )

        # --- コーナーハンドル (●) ---
        r = self.corner_radius
        for i, (cx, cy) in enumerate(c_pts_tuples):
            self.canvas.create_oval(
                cx - r,
                cy - r,
                cx + r,
                cy + r,
                fill="#00E5FF",
                outline="white",
                width=2,
                tags=("ui", "handle", f"corner_{i}"),
            )

        # --- エッジハンドル (■) ---
        er = self.edge_radius
        for i in range(4):
            p1 = c_pts_tuples[i]
            p2 = c_pts_tuples[(i + 1) % 4]

            # 中点
            mx = (p1[0] + p2[0]) / 2
            my = (p1[1] + p2[1]) / 2

            self.canvas.create_rectangle(
                mx - er,
                my - er,
                mx + er,
                my + er,
                fill="#FF4081",
                outline="white",
                width=2,
                tags=("ui", "handle", f"edge_{i}"),
            )

    # --- 操作イベントハンドラ ---

    def _on_wheel(self, event):
        """マウスホイールでズーム"""
        if event.num == 5 or event.delta < 0:
            factor = 0.9
        else:
            factor = 1.1

        new_zoom = self.zoom_level * factor

        if 0.1 < new_zoom < 10.0:
            imx, imy = self._canvas_to_img(event.x, event.y)
            self.zoom_level = new_zoom
            new_cx, new_cy = self._img_to_canvas(imx, imy)
            self.pan_x += event.x - new_cx
            self.pan_y += event.y - new_cy
            self.display_image = None
            self._redraw_all()

    def _on_pan_start(self, event):
        self.canvas.config(cursor="fleur")
        self.drag_state["type"] = "pan"
        self.drag_state["last_x"] = event.x
        self.drag_state["last_y"] = event.y

    def _on_pan_drag(self, event):
        if self.drag_state["type"] == "pan":
            dx = event.x - self.drag_state["last_x"]
            dy = event.y - self.drag_state["last_y"]
            self.pan_x += dx
            self.pan_y += dy
            self.drag_state["last_x"] = event.x
            self.drag_state["last_y"] = event.y
            self._redraw_all()

    def _on_left_down(self, event):
        """左クリック: ハンドルを掴む"""
        items = self.canvas.find_overlapping(
            event.x - 8, event.y - 8, event.x + 8, event.y + 8
        )

        target_type = None
        target_id = None

        for item in items:
            tags = self.canvas.gettags(item)
            if "handle" in tags:
                for tag in tags:
                    if tag.startswith("corner_"):
                        target_type = "corner"
                        target_id = int(tag.split("_")[1])
                        break
                    elif tag.startswith("edge_"):
                        target_type = "edge"
                        target_id = int(tag.split("_")[1])
                        break
            if target_type == "corner":
                break

        if target_type:
            self.drag_state["type"] = target_type
            self.drag_state["item"] = target_id
            self.drag_state["last_x"] = event.x
            self.drag_state["last_y"] = event.y

            if target_type == "corner":
                self._update_magnifier(target_id, event.x, event.y)

    def _on_left_up(self, event):
        self.drag_state["type"] = None
        self.drag_state["item"] = None
        self.canvas.delete("magnifier")
        self.canvas.config(cursor="crosshair")

    def _on_left_drag(self, event):
        dtype = self.drag_state["type"]
        idx = self.drag_state["item"]

        if dtype is None:
            return

        iw, ih = self.original_image.size

        if dtype == "corner":
            # --- 角の移動 ---
            ix, iy = self._canvas_to_img(event.x, event.y)
            ix = max(0, min(iw, ix))
            iy = max(0, min(ih, iy))

            self.corners[idx] = [ix, iy]
            self._update_magnifier(idx, event.x, event.y)

        elif dtype == "edge":
            # --- 辺の平行移動 (軸固定) ---
            dx_px = event.x - self.drag_state["last_x"]
            dy_px = event.y - self.drag_state["last_y"]

            dx_img = dx_px / self.current_scale
            dy_img = dy_px / self.current_scale

            # ★ 移動方向の制限
            # edge 0 (Top) / 2 (Bottom) -> 縦移動のみ (dx=0)
            # edge 1 (Right) / 3 (Left) -> 横移動のみ (dy=0)
            if idx % 2 == 0:
                dx_img = 0
            else:
                dy_img = 0

            p1_idx = idx
            p2_idx = (idx + 1) % 4

            p1 = self.corners[p1_idx]
            p2 = self.corners[p2_idx]

            # 移動後の座標候補
            n1x, n1y = p1[0] + dx_img, p1[1] + dy_img
            n2x, n2y = p2[0] + dx_img, p2[1] + dy_img

            # 範囲制限
            if n1x < 0:
                dx_img += 0 - n1x
            if n1x > iw:
                dx_img += iw - n1x
            if n2x < 0:
                dx_img += 0 - n2x
            if n2x > iw:
                dx_img += iw - n2x

            if n1y < 0:
                dy_img += 0 - n1y
            if n1y > ih:
                dy_img += ih - n1y
            if n2y < 0:
                dy_img += 0 - n2y
            if n2y > ih:
                dy_img += ih - n2y

            self.corners[p1_idx][0] += dx_img
            self.corners[p1_idx][1] += dy_img
            self.corners[p2_idx][0] += dx_img
            self.corners[p2_idx][1] += dy_img

            self.drag_state["last_x"] = event.x
            self.drag_state["last_y"] = event.y

        self._draw_overlays()

    def _update_magnifier(self, corner_idx, canvas_x, canvas_y):
        """拡大鏡の描画"""
        self.canvas.delete("magnifier")

        img_pos = self.corners[corner_idx]
        img_x, img_y = img_pos[0], img_pos[1]

        crop_w = self.magnifier_size / self.magnifier_zoom
        crop_h = self.magnifier_size / self.magnifier_zoom

        left = img_x - crop_w / 2
        top = img_y - crop_h / 2
        right = img_x + crop_w / 2
        bottom = img_y + crop_h / 2

        try:
            cropped = self.original_image.crop((left, top, right, bottom))
        except Exception:
            return

        resized = cropped.resize(
            (self.magnifier_size, self.magnifier_size), Image.Resampling.NEAREST
        )
        self.magnifier_image_ref = ImageTk.PhotoImage(resized)

        offset = 30
        mag_x = canvas_x + offset
        mag_y = canvas_y + offset

        if mag_x + self.magnifier_size > self.canvas.winfo_width():
            mag_x = canvas_x - self.magnifier_size - offset
        if mag_y + self.magnifier_size > self.canvas.winfo_height():
            mag_y = canvas_y - self.magnifier_size - offset

        self.canvas.create_rectangle(
            mag_x,
            mag_y,
            mag_x + self.magnifier_size,
            mag_y + self.magnifier_size,
            outline="#00E5FF",
            width=3,
            fill="black",
            tags="magnifier",
        )
        self.canvas.create_image(
            mag_x, mag_y, anchor="nw", image=self.magnifier_image_ref, tags="magnifier"
        )
        cx = mag_x + self.magnifier_size / 2
        cy = mag_y + self.magnifier_size / 2
        L = 10
        self.canvas.create_line(
            cx - L, cy, cx + L, cy, fill="#00E5FF", width=1, tags="magnifier"
        )
        self.canvas.create_line(
            cx, cy - L, cx, cy + L, fill="#00E5FF", width=1, tags="magnifier"
        )

    def _reset_corners(self):
        iw, ih = self.original_image.size
        self.corners = [[0, 0], [iw, 0], [iw, ih], [0, ih]]
        self._draw_overlays()

    def _apply_and_save(self):
        """射影変換を実行してコールバックを呼ぶ"""
        try:
            transformed_img = self._perspective_transform(
                self.original_image, self.corners
            )
            self.on_save_callback(transformed_img)
            self.destroy()
        except Exception as e:
            tk.messagebox.showerror("エラー", f"変換に失敗しました: {e}", parent=self)

    def _perspective_transform(self, img, src_points):
        pts = np.array(src_points, dtype="float32")
        (tl, tr, br, bl) = pts

        widthA = np.sqrt(((br[0] - bl[0]) ** 2) + ((br[1] - bl[1]) ** 2))
        widthB = np.sqrt(((tr[0] - tl[0]) ** 2) + ((tr[1] - tl[1]) ** 2))
        maxWidth = max(int(widthA), int(widthB))

        heightA = np.sqrt(((tr[0] - br[0]) ** 2) + ((tr[1] - br[1]) ** 2))
        heightB = np.sqrt(((tl[0] - bl[0]) ** 2) + ((tl[1] - bl[1]) ** 2))
        maxHeight = max(int(heightA), int(heightB))

        dst = np.array(
            [
                [0, 0],
                [maxWidth - 1, 0],
                [maxWidth - 1, maxHeight - 1],
                [0, maxHeight - 1],
            ],
            dtype="float32",
        )

        matrix = self._get_perspective_transform_matrix(dst, pts)

        if matrix is None:
            return img

        matrix = matrix / matrix[2, 2]
        coeffs = matrix.flatten()[:8]

        return img.transform(
            (maxWidth, maxHeight), Image.PERSPECTIVE, coeffs, Image.Resampling.BICUBIC
        )

    def _get_perspective_transform_matrix(self, src, dst):
        matrix = []
        for (x, y), (X, Y) in zip(src, dst):
            matrix.extend(
                [[x, y, 1, 0, 0, 0, -X * x, -X * y], [0, 0, 0, x, y, 1, -Y * x, -Y * y]]
            )

        A = np.array(matrix, dtype=float)
        B = np.array(dst, dtype=float).reshape(8)

        try:
            res = np.linalg.solve(A, B)
            H = np.append(res, 1).reshape(3, 3)
            return H
        except np.linalg.LinAlgError:
            return None

    def iconbitmap(self, *args, **kwargs):
        """
        CustomTkinterがアイコンをリセットするのを防ぐためのオーバーライドメソッド。
        """
        if self._custom_icon_path:
            try:
                super().iconbitmap(self._custom_icon_path)
            except Exception:
                pass
        else:
            try:
                super().iconbitmap(*args, **kwargs)
            except Exception:
                pass

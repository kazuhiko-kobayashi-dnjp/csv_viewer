"""軽量CSVビューワ (Canvas仮想スクロール版)

特徴:
- 複数ファイルをタブで表示
- 行番号列 + 列ヘッダ表示
- Excelライクな矩形セル選択（ドラッグ / Shift+矢印 / Ctrl+Shift+矢印で端まで一括）
- ヘッダ(ラベル名)も選択してコピー可能
- 2段ヘッダ: 接頭語(ドット前)を共通列でまとめて上段に、個別名(ドット後)を下段に表示
- 巨大CSV対応: 行オフセットをバックグラウンドで索引し、表示中の行だけ読む仮想スクロール
- 先頭行(ヘッダ)と先頭列(行番号)を固定
"""

import os
import sys
import csv
import shutil
import subprocess
import threading
import tkinter as tk
from tkinter import ttk, filedialog, messagebox

ROW_H = 20            # データ行・ヘッダ各段の高さ(px)
ROWNUM_W = 70         # 行番号列の幅(px)
COL_MIN_W = 60
COL_MAX_W = 280
CHAR_W = 7            # 文字幅見積り(px)
FONT = ("TkDefaultFont", 9)
CACHE_LIMIT = 4000    # 行キャッシュ上限


def _make_tk():
    try:
        from tkinterdnd2 import TkinterDnD
        return TkinterDnD.Tk()
    except Exception:
        return tk.Tk()


def split_label(name):
    """ラベルをドットで全階層分割してリストで返す。"""
    return name.split(".")


def to_local_path(path):
    """ドロップされたパスを実行環境に合わせて正規化する。

    - ネイティブWindows: そのまま使う。
    - WSL/Linux: Windowsパス(C:\\...)を /mnt/c/... へ変換。
    """
    path = path.strip().strip("{}").strip('"')
    if os.name == "nt":
        return path
    if len(path) >= 2 and path[1] == ":" and path[0].isalpha():
        wslpath = shutil.which("wslpath")
        if wslpath:
            try:
                return subprocess.check_output(
                    [wslpath, "-u", path], text=True).strip()
            except Exception:
                pass
        drive = path[0].lower()
        rest = path[2:].replace("\\", "/").lstrip("/")
        return f"/mnt/{drive}/{rest}"
    return path


def parse_drop(data):
    """tkinterdnd2 の Drop データから複数パスを取り出す。"""
    paths = []
    s = data.strip()
    i = 0
    while i < len(s):
        if s[i] == "{":
            j = s.find("}", i)
            if j == -1:
                paths.append(s[i + 1:])
                break
            paths.append(s[i + 1:j])
            i = j + 1
            while i < len(s) and s[i] == " ":
                i += 1
        else:
            j = s.find(" ", i)
            if j == -1:
                paths.append(s[i:])
                break
            paths.append(s[i:j])
            i = j + 1
    return [to_local_path(p) for p in paths if p.strip()]


class Grid(ttk.Frame):
    """1ファイル分の表示・操作を担当するウィジェット。"""

    def __init__(self, parent, set_status, open_in_new_tab):
        super().__init__(parent)
        self.set_status = set_status
        self.open_in_new_tab = open_in_new_tab

        self._filepath = None
        self._encoding = "utf-8"
        self._headers = []
        self._parts = []      # 各列: ドット分割された階層リスト
        self._header_rows = 1
        self._col_w = []
        self._col_x = []
        self._total_w = 0

        self._offsets = [0]
        self._total_rows = 0
        self._indexing = False
        self._pending = None
        self._index_error = None

        self._row_cache = {}
        self._cache_order = []

        self._first_row = 0
        self._x_off = 0

        # セル選択 (r0,c0)=アンカー (r1,c1)=アクティブ端
        self._sel = None
        # ヘッダ(ラベル)選択 (c0, c1)
        self._hsel = None
        self._dragging = None      # "cell" | "header" | None

        self._search_result = None
        self._search_term = ""

        self._resize_col = None   # リサイズ中の列インデックス
        self._resize_x0 = 0      # ドラッグ開始X
        self._resize_w0 = 0      # ドラッグ開始時の列幅

        self._tip_win = None     # ツールチップ Toplevel
        self._tip_col = -1       # 現在表示中のツールチップ列

        self._build_ui()
        self._setup_dnd()

    # ── UI ────────────────────────────────────────────────────────
    def _build_ui(self):
        self._canvas = tk.Canvas(self, bg="white", highlightthickness=0)
        self._vsb = ttk.Scrollbar(self, orient=tk.VERTICAL, command=self._on_vscroll)
        self._hsb = ttk.Scrollbar(self, orient=tk.HORIZONTAL, command=self._on_hscroll)
        self._canvas.grid(row=0, column=0, sticky="nsew")
        self._vsb.grid(row=0, column=1, sticky="ns")
        self._hsb.grid(row=1, column=0, sticky="ew")
        self.rowconfigure(0, weight=1)
        self.columnconfigure(0, weight=1)

        cv = self._canvas
        cv.bind("<Configure>", lambda e: self._redraw())
        cv.bind("<MouseWheel>", self._on_wheel)
        cv.bind("<Button-4>", self._on_wheel)
        cv.bind("<Button-5>", self._on_wheel)
        cv.bind("<Shift-MouseWheel>", self._on_wheel_h)
        cv.bind("<Button-1>", self._on_press)
        cv.bind("<B1-Motion>", self._on_drag)
        cv.bind("<ButtonRelease-1>", self._on_release)
        cv.bind("<Control-c>", lambda e: self.copy_selection())
        cv.bind("<Control-C>", lambda e: self.copy_selection())
        # 矢印: 単独移動 / Ctrl: 端まで移動 / Shift: 範囲拡張 / Ctrl+Shift: 端まで拡張
        for key, dr, dc in (("Up", -1, 0), ("Down", 1, 0),
                            ("Left", 0, -1), ("Right", 0, 1)):
            cv.bind(f"<{key}>",               lambda e, a=dr, b=dc: self._move_sel(a, b, False))
            cv.bind(f"<Control-{key}>",       lambda e, a=dr, b=dc: self._move_sel(a, b, True))
            cv.bind(f"<Shift-{key}>",         lambda e, a=dr, b=dc: self._extend_sel(a, b, False))
            cv.bind(f"<Control-Shift-{key}>", lambda e, a=dr, b=dc: self._extend_sel(a, b, True))
        cv.bind("<Prior>",          lambda e: self._page(-1))
        cv.bind("<Next>",           lambda e: self._page(1))
        cv.bind("<Shift-Prior>",    lambda e: self._page_extend(-1))
        cv.bind("<Shift-Next>",     lambda e: self._page_extend(1))
        cv.bind("<Home>",           lambda e: self._move_home())
        cv.bind("<End>",            lambda e: self._move_end())
        cv.bind("<Shift-Home>",     lambda e: self._extend_home())
        cv.bind("<Shift-End>",      lambda e: self._extend_end())
        cv.bind("<Control-Home>",       lambda _e: self._move_corner(0, 0))
        cv.bind("<Control-End>",        lambda _e: self._move_corner(self._total_rows - 1, len(self._headers) - 1))
        cv.bind("<Control-Shift-Home>", lambda _e: self._extend_corner(0, 0))
        cv.bind("<Control-Shift-End>",  lambda _e: self._extend_corner(self._total_rows - 1, len(self._headers) - 1))
        cv.bind("<Control-a>",      lambda e: self._select_all())
        cv.bind("<Control-A>",      lambda e: self._select_all())
        cv.bind("<Motion>",         self._on_motion)
        cv.bind("<Leave>",          lambda e: self._hide_tooltip())

    def _setup_dnd(self):
        try:
            from tkinterdnd2 import DND_FILES
            self._canvas.drop_target_register(DND_FILES)
            self._canvas.dnd_bind("<<Drop>>", self._on_drop)
        except Exception:
            pass

    def _on_drop(self, event):
        for path in parse_drop(event.data):
            self.open_in_new_tab(path)

    def focus_grid(self):
        self._canvas.focus_set()

    # ── ファイル読込 ─────────────────────────────────────────────
    def open_file(self, path):
        try:
            enc = self._detect_encoding(path)
            with open(path, "r", encoding=enc, newline="") as f:
                header_line = f.readline()
            if not header_line:
                messagebox.showerror("Error", "空のファイルです")
                return False
            headers = next(csv.reader([header_line]))
        except Exception as e:
            messagebox.showerror("Error", f"読込失敗: {e}")
            return False

        self._filepath = path
        self._encoding = enc
        self._headers = headers
        self._parts = [split_label(h) for h in headers]
        self._header_rows = max(len(p) for p in self._parts) if self._parts else 1
        self._compute_columns()

        self._row_cache.clear()
        self._cache_order.clear()
        self._offsets = []
        self._total_rows = 0
        self._first_row = 0
        self._x_off = 0
        self._sel = None
        self._hsel = None

        self._set_status(f"{path}  ({len(headers)} cols)  索引中…")
        self._start_indexing(path, enc, len(header_line.encode(enc)))
        return True

    def _set_status(self, text):
        # 自分が現在のタブのときのみステータス更新
        self.set_status(self, text)

    def _detect_encoding(self, path):
        with open(path, "rb") as f:
            head = f.read(4)
        if head.startswith(b"\xef\xbb\xbf"):
            return "utf-8-sig"
        for enc in ("utf-8", "cp932"):
            try:
                with open(path, "r", encoding=enc) as f:
                    f.read(65536)
                return enc
            except UnicodeDecodeError:
                continue
        return "utf-8"

    def _compute_columns(self):
        self._col_w = []
        self._col_x = []
        x = 0
        for parts in self._parts:
            label_len = max(len(s) for s in parts)
            w = max(COL_MIN_W, min(COL_MAX_W, label_len * CHAR_W + 12))
            self._col_w.append(w)
            self._col_x.append(x)
            x += w
        self._total_w = x

    def _start_indexing(self, path, enc, header_bytes):
        self._indexing = True
        self._pending = None
        self._index_error = None

        def worker():
            try:
                offsets = [header_bytes]
                with open(path, "rb") as f:
                    f.seek(header_bytes)
                    pos = header_bytes
                    for line in f:
                        pos += len(line)
                        offsets.append(pos)
                        if len(offsets) % 50000 == 0:
                            self._pending = (list(offsets[:-1]), False)
                self._pending = (offsets[:-1], True)
            except Exception as e:
                self._index_error = str(e)

        threading.Thread(target=worker, daemon=True).start()
        self.after(100, self._poll_index)

    def _poll_index(self):
        if self._index_error is not None:
            self._set_status(f"索引エラー: {self._index_error}")
            self._indexing = False
            return
        if self._pending is not None:
            offsets, done = self._pending
            self._pending = None
            self._offsets = offsets
            self._total_rows = len(offsets)
            if done:
                self._indexing = False
                self._set_status(
                    f"{self._filepath}  {self._total_rows} rows × {len(self._headers)} cols")
                self._redraw()
                return
            self._set_status(f"{self._filepath}  索引中… {self._total_rows} rows")
            self._redraw()
        if self._indexing:
            self.after(100, self._poll_index)

    # ── 行データ取得 ─────────────────────────────────────────────
    def _get_row(self, idx):
        if idx in self._row_cache:
            return self._row_cache[idx]
        if idx >= self._total_rows or not self._filepath:
            return None
        try:
            start = self._offsets[idx]
            with open(self._filepath, "r", encoding=self._encoding, newline="") as f:
                f.seek(start)
                line = f.readline()
            fields = next(csv.reader([line])) if line else []
        except Exception:
            fields = []
        self._row_cache[idx] = fields
        self._cache_order.append(idx)
        if len(self._cache_order) > CACHE_LIMIT:
            old = self._cache_order.pop(0)
            self._row_cache.pop(old, None)
        return fields

    # ── スクロール ───────────────────────────────────────────────
    def _visible_data_rows(self):
        h = self._canvas.winfo_height()
        return max(1, (h - self._header_rows * ROW_H) // ROW_H)

    def _on_vscroll(self, *args):
        if not self._total_rows:
            return
        if args[0] == "moveto":
            self._first_row = int(float(args[1]) * self._total_rows)
        elif args[0] == "scroll":
            n = int(args[1])
            if args[2] == "pages":
                n *= self._visible_data_rows()
            self._first_row += n
        self._clamp_scroll()
        self._redraw()

    def _on_hscroll(self, *args):
        if args[0] == "moveto":
            self._x_off = int(float(args[1]) * self._total_w)
        elif args[0] == "scroll":
            self._x_off += int(args[1]) * 40
        self._clamp_scroll()
        self._redraw()

    def _on_wheel(self, event):
        if getattr(event, "num", None) == 4:
            delta = -3
        elif getattr(event, "num", None) == 5:
            delta = 3
        else:
            delta = -int(event.delta / 120) * 3
        self._first_row += delta
        self._clamp_scroll()
        self._redraw()

    def _on_wheel_h(self, event):
        self._x_off += -int(event.delta / 120) * 40
        self._clamp_scroll()
        self._redraw()

    def _clamp_scroll(self):
        max_row = max(0, self._total_rows - self._visible_data_rows())
        self._first_row = max(0, min(self._first_row, max_row))
        w = self._canvas.winfo_width() - ROWNUM_W
        max_x = max(0, self._total_w - w)
        self._x_off = max(0, min(self._x_off, max_x))

    # ── 座標変換 ─────────────────────────────────────────────────
    def _col_at_x(self, px):
        if px < ROWNUM_W:
            return None
        data_x = px - ROWNUM_W + self._x_off
        if data_x < 0 or data_x >= self._total_w:
            return None
        lo, hi = 0, len(self._col_x) - 1
        while lo <= hi:
            mid = (lo + hi) // 2
            x0 = self._col_x[mid]
            x1 = x0 + self._col_w[mid]
            if data_x < x0:
                hi = mid - 1
            elif data_x >= x1:
                lo = mid + 1
            else:
                return mid
        return None

    def _row_at_y(self, py):
        top = self._header_rows * ROW_H
        if py < top:
            return None
        r = self._first_row + (py - top) // ROW_H
        if r >= self._total_rows:
            return None
        return int(r)

    # ── リサイズ境界検出 ─────────────────────────────────────────
    def _resize_col_at(self, px):
        """px 付近（±4px）にある列境界の列インデックスを返す。なければ None。"""
        SNAP = 4
        for i, (cx, cw) in enumerate(zip(self._col_x, self._col_w)):
            rx = ROWNUM_W + cx + cw - self._x_off
            if abs(px - rx) <= SNAP:
                return i
        return None

    # ── 選択 ─────────────────────────────────────────────────────
    def _on_press(self, event):
        self._canvas.focus_set()
        self._hide_tooltip()
        if event.y < self._header_rows * ROW_H and self._headers:
            rc = self._resize_col_at(event.x)
            if rc is not None:
                self._dragging = "resize_col"
                self._resize_col = rc
                self._resize_x0 = event.x
                self._resize_w0 = self._col_w[rc]
                return
            c = self._col_at_x(event.x)
            if c is None:
                return
            self._hsel = (c, c)
            self._sel = None
            self._dragging = "header"
            self._redraw()
            return
        c = self._col_at_x(event.x)
        r = self._row_at_y(event.y)
        if c is None or r is None:
            return
        self._sel = (r, c, r, c)
        self._hsel = None
        self._dragging = "cell"
        self._update_cell_status()
        self._redraw()

    def _on_drag(self, event):
        if self._dragging == "resize_col":
            new_w = max(COL_MIN_W, self._resize_w0 + event.x - self._resize_x0)
            self._col_w[self._resize_col] = new_w
            # col_x を再計算
            x = 0
            for i in range(len(self._col_x)):
                self._col_x[i] = x
                x += self._col_w[i]
            self._total_w = x
            self._clamp_scroll()
            self._redraw()
            return
        if self._dragging == "header":
            cx = max(ROWNUM_W, min(event.x, self._canvas.winfo_width() - 1))
            c = self._col_at_x(cx)
            if c is None:
                c = self._hsel[1]
            self._hsel = (self._hsel[0], c)
            self._redraw()
            return
        if self._dragging != "cell" or self._sel is None:
            return
        if event.y > self._canvas.winfo_height() - ROW_H:
            self._first_row += 1
        elif event.y < self._header_rows * ROW_H:
            self._first_row -= 1
        self._clamp_scroll()
        cx = max(ROWNUM_W, min(event.x, self._canvas.winfo_width() - 1))
        cy = max(self._header_rows * ROW_H, min(event.y, self._canvas.winfo_height() - 1))
        c = self._col_at_x(cx)
        r = self._row_at_y(cy)
        if c is None:
            c = self._sel[3]
        if r is None:
            r = self._sel[2]
        self._sel = (self._sel[0], self._sel[1], r, c)
        self._redraw()

    def _on_release(self, event):
        self._dragging = None
        self._resize_col = None

    def _on_motion(self, event):
        if not self._headers:
            return
        # リサイズカーソル切り替え
        if event.y < self._header_rows * ROW_H:
            if self._resize_col_at(event.x) is not None:
                self._canvas.config(cursor="sb_h_double_arrow")
            else:
                self._canvas.config(cursor="")
            # ツールチップ
            c = self._col_at_x(event.x)
            if c is not None:
                self._show_tooltip(event, c)
            else:
                self._hide_tooltip()
        else:
            self._canvas.config(cursor="")
            self._hide_tooltip()

    def _show_tooltip(self, event, col):
        if self._tip_col == col:
            # 位置だけ更新
            if self._tip_win:
                x = self._canvas.winfo_rootx() + event.x + 12
                y = self._canvas.winfo_rooty() + event.y + 16
                self._tip_win.geometry(f"+{x}+{y}")
            return
        self._hide_tooltip()
        self._tip_col = col
        label = self._headers[col]
        x = self._canvas.winfo_rootx() + event.x + 12
        y = self._canvas.winfo_rooty() + event.y + 16
        win = tk.Toplevel(self._canvas)
        win.wm_overrideredirect(True)
        win.geometry(f"+{x}+{y}")
        tk.Label(win, text=label, background="#ffffe0", relief="solid",
                 borderwidth=1, font=FONT, padx=4, pady=2).pack()
        self._tip_win = win

    def _hide_tooltip(self):
        if self._tip_win:
            self._tip_win.destroy()
            self._tip_win = None
        self._tip_col = -1

    def _norm_sel(self):
        r0, c0, r1, c1 = self._sel
        return min(r0, r1), min(c0, c1), max(r0, r1), max(c0, c1)

    def copy_selection(self):
        # ヘッダ選択が優先（ラベル名コピー）
        if self._hsel is not None:
            c0, c1 = min(self._hsel), max(self._hsel)
            labels = [self._headers[c] for c in range(c0, c1 + 1)]
            self._to_clipboard("\t".join(labels))
            self._set_status(f"ラベル {c1 - c0 + 1} 列をコピーしました")
            return "break"
        if not self._sel:
            return "break"
        r0, c0, r1, c1 = self._norm_sel()
        lines = []
        for r in range(r0, r1 + 1):
            row = self._get_row(r) or []
            cells = [row[c] if c < len(row) else "" for c in range(c0, c1 + 1)]
            lines.append("\t".join(cells))
        self._to_clipboard("\n".join(lines))
        self._set_status(f"{r1 - r0 + 1}行 × {c1 - c0 + 1}列をコピーしました")
        return "break"

    def copy_with_labels(self):
        """選択範囲を、対象列のラベル名を1行目に付けてコピー。"""
        if self._hsel is not None:
            return self.copy_selection()
        if not self._sel:
            return "break"
        r0, c0, r1, c1 = self._norm_sel()
        lines = ["\t".join(self._headers[c] for c in range(c0, c1 + 1))]
        for r in range(r0, r1 + 1):
            row = self._get_row(r) or []
            lines.append("\t".join(row[c] if c < len(row) else "" for c in range(c0, c1 + 1)))
        self._to_clipboard("\n".join(lines))
        self._set_status(f"ラベル付きで {r1 - r0 + 1}行をコピーしました")
        return "break"

    def _to_clipboard(self, text):
        self.clipboard_clear()
        self.clipboard_append(text)

    def get_plot_info(self):
        """グラフウィンドウ起動に必要な情報を返す（データは読まない）。
        戻り値: {
            'headers': [全列名],
            'sel_cols': [選択列インデックス],
            'total_rows': int,
            'filepath': str, 'encoding': str, 'offsets': list,
        }
        """
        if not self._headers or not self._total_rows:
            return None
        if self._hsel is not None:
            c0, c1 = min(self._hsel), max(self._hsel)
            sel_cols = list(range(c0, c1 + 1))
        elif self._sel is not None:
            _, c0, _, c1 = self._sel
            c0, c1 = min(c0, c1), max(c0, c1)
            sel_cols = list(range(c0, c1 + 1))
        else:
            return None
        return {
            "headers": self._headers,
            "sel_cols": sel_cols,
            "total_rows": self._total_rows,
            "filepath": self._filepath,
            "encoding": self._encoding,
            "offsets": self._offsets,
        }

    # ── キーボード移動 ───────────────────────────────────────────
    def _move_sel(self, dr, dc, jump=False):
        if not self._headers or not self._total_rows:
            return "break"
        if self._sel is None:
            r, c = self._first_row, 0
        else:
            _, _, r, c = self._sel
        if jump:
            if dr < 0: r = 0
            elif dr > 0: r = self._total_rows - 1
            if dc < 0: c = 0
            elif dc > 0: c = len(self._headers) - 1
        else:
            r = max(0, min(r + dr, self._total_rows - 1))
            c = max(0, min(c + dc, len(self._headers) - 1))
        self._sel = (r, c, r, c)
        self._hsel = None
        self._ensure_visible(r, c)
        self._update_cell_status()
        self._redraw()
        return "break"

    def _extend_sel(self, dr, dc, jump=False):
        if not self._headers or not self._total_rows:
            return "break"
        if self._sel is None:
            self._sel = (self._first_row, 0, self._first_row, 0)
        r0, c0, r1, c1 = self._sel
        if jump:
            if dr < 0: r1 = 0
            elif dr > 0: r1 = self._total_rows - 1
            if dc < 0: c1 = 0
            elif dc > 0: c1 = len(self._headers) - 1
        else:
            r1 = max(0, min(r1 + dr, self._total_rows - 1))
            c1 = max(0, min(c1 + dc, len(self._headers) - 1))
        self._sel = (r0, c0, r1, c1)
        self._hsel = None
        self._ensure_visible(r1, c1)
        self._redraw()
        return "break"

    def _page(self, direction):
        n = self._visible_data_rows()
        if self._sel:
            _, c0, _, c1 = self._sel
            r = max(0, min(self._sel[2] + direction * n, self._total_rows - 1))
            self._sel = (r, c0, r, c1)
            self._ensure_visible(r, c1)
            self._update_cell_status()
        else:
            self._first_row += direction * n
        self._clamp_scroll()
        self._redraw()
        return "break"

    def _page_extend(self, direction):
        if not self._headers or not self._total_rows:
            return "break"
        if self._sel is None:
            self._sel = (self._first_row, 0, self._first_row, 0)
        r0, c0, r1, c1 = self._sel
        n = self._visible_data_rows()
        r1 = max(0, min(r1 + direction * n, self._total_rows - 1))
        self._sel = (r0, c0, r1, c1)
        self._hsel = None
        self._ensure_visible(r1, c1)
        self._redraw()
        return "break"

    def _move_home(self):
        if not self._headers or not self._total_rows:
            return "break"
        r = self._sel[2] if self._sel else self._first_row
        self._sel = (r, 0, r, 0)
        self._hsel = None
        self._ensure_visible(r, 0)
        self._update_cell_status()
        self._redraw()
        return "break"

    def _move_end(self):
        if not self._headers or not self._total_rows:
            return "break"
        r = self._sel[2] if self._sel else self._first_row
        c = len(self._headers) - 1
        self._sel = (r, c, r, c)
        self._hsel = None
        self._ensure_visible(r, c)
        self._update_cell_status()
        self._redraw()
        return "break"

    def _extend_home(self):
        if not self._headers or not self._total_rows:
            return "break"
        if self._sel is None:
            self._sel = (self._first_row, 0, self._first_row, 0)
        r0, c0, r1, _ = self._sel
        self._sel = (r0, c0, r1, 0)
        self._hsel = None
        self._ensure_visible(r1, 0)
        self._redraw()
        return "break"

    def _extend_end(self):
        if not self._headers or not self._total_rows:
            return "break"
        if self._sel is None:
            self._sel = (self._first_row, 0, self._first_row, 0)
        r0, c0, r1, _ = self._sel
        c = len(self._headers) - 1
        self._sel = (r0, c0, r1, c)
        self._hsel = None
        self._ensure_visible(r1, c)
        self._redraw()
        return "break"

    def _select_all(self):
        if not self._headers or not self._total_rows:
            return "break"
        self._sel = (0, 0, self._total_rows - 1, len(self._headers) - 1)
        self._hsel = None
        self._redraw()
        return "break"

    def _move_corner(self, r, c):
        if not self._headers or not self._total_rows:
            return "break"
        r = max(0, min(r, self._total_rows - 1))
        c = max(0, min(c, len(self._headers) - 1))
        self._sel = (r, c, r, c)
        self._hsel = None
        self._ensure_visible(r, c)
        self._update_cell_status()
        self._redraw()
        return "break"

    def _extend_corner(self, r, c):
        if not self._headers or not self._total_rows:
            return "break"
        if self._sel is None:
            self._sel = (self._first_row, 0, self._first_row, 0)
        r0, c0 = self._sel[0], self._sel[1]
        r = max(0, min(r, self._total_rows - 1))
        c = max(0, min(c, len(self._headers) - 1))
        self._sel = (r0, c0, r, c)
        self._hsel = None
        self._ensure_visible(r, c)
        self._redraw()
        return "break"

    def _ensure_visible(self, r, c):
        n_vis = self._visible_data_rows()
        if r < self._first_row:
            self._first_row = r
        elif r >= self._first_row + n_vis:
            self._first_row = r - n_vis + 1
        x0 = self._col_x[c]
        x1 = x0 + self._col_w[c]
        avail = self._canvas.winfo_width() - ROWNUM_W
        if x0 < self._x_off:
            self._x_off = x0
        elif x1 > self._x_off + avail:
            self._x_off = x1 - avail
        self._clamp_scroll()

    # ── 検索 ─────────────────────────────────────────────────────
    def search(self, term, direction):
        term = term.strip()
        if not term or not self._total_rows:
            return
        terms = [t.lower() for t in term.split() if t]
        if self._sel:
            _, _, sr, sc = self._sel
        else:
            sr, sc = self._first_row, -1
        ncol = len(self._headers)
        path = self._filepath
        enc = self._encoding
        offsets = self._offsets
        total = self._total_rows
        header_lower = [h.lower() for h in self._headers]

        def read_row(i):
            try:
                with open(path, "r", encoding=enc, newline="") as f:
                    f.seek(offsets[i])
                    line = f.readline()
                return next(csv.reader([line])) if line else []
            except Exception:
                return []

        def thread():
            r, c = sr, sc
            rows_scanned = 0
            single = len(terms) == 1
            while rows_scanned <= total:
                row = read_row(r)
                lowered = [v.lower() for v in row]

                def text_at(cc):
                    cell = lowered[cc] if cc < len(lowered) else ""
                    label = header_lower[cc] if cc < len(header_lower) else ""
                    return cell + " " + label

                if single:
                    cols = range(c + 1, ncol) if direction > 0 else range(c - 1, -1, -1)
                    for cc in cols:
                        if terms[0] in text_at(cc):
                            self._search_result = (r, cc)
                            return
                else:
                    texts = [text_at(cc) for cc in range(ncol)]
                    if all(any(t in tx for tx in texts) for t in terms):
                        for cc in range(ncol):
                            if any(t in texts[cc] for t in terms):
                                self._search_result = (r, cc)
                                return
                r += direction
                if r < 0:
                    r = total - 1
                elif r >= total:
                    r = 0
                c = -1 if direction > 0 else ncol
                rows_scanned += 1
            self._search_result = "notfound"

        self._search_result = None
        self._search_term = term
        self._set_status(f"検索中: {term}")
        threading.Thread(target=thread, daemon=True).start()
        self.after(50, self._poll_search)

    def _poll_search(self):
        res = self._search_result
        if res is None:
            self.after(50, self._poll_search)
            return
        if res == "notfound":
            self._set_status(f"「{self._search_term}」は見つかりません")
            return
        self._goto_match(res[0], res[1])

    def _goto_match(self, r, c):
        self._sel = (r, c, r, c)
        self._hsel = None
        self._ensure_visible(r, c)
        self._set_status(f"一致: 行{r + 1} 列{c + 1} ({self._headers[c]})")
        self._canvas.focus_set()
        self._redraw()

    def _update_cell_status(self):
        if self._sel is None or not self._headers:
            return
        _, _, r, c = self._sel
        col_name = self._headers[c] if c < len(self._headers) else ""
        self._set_status(f"{self._filepath}  行{r + 1} 列{c + 1} ({col_name})")

    # ── 描画 ─────────────────────────────────────────────────────
    def _redraw(self):
        cv = self._canvas
        cv.delete("all")
        if not self._headers:
            cv.create_text(20, 20, anchor="nw",
                           text="CSVファイルを開いてください（Open CSV / ドラッグ&ドロップ）",
                           fill="gray", font=FONT)
            return

        W = cv.winfo_width()
        H = cv.winfo_height()
        n_vis = self._visible_data_rows()
        header_h = self._header_rows * ROW_H
        data_top = header_h

        first_col, last_col = self._visible_col_range(W)

        if self._sel:
            self._draw_crosshair(cv, first_col, last_col, n_vis, data_top, W, H)
            self._draw_selection(cv, first_col, last_col, n_vis, data_top)

        for vi in range(n_vis):
            r = self._first_row + vi
            if r >= self._total_rows:
                break
            y = data_top + vi * ROW_H
            row = self._get_row(r)
            for c in range(first_col, last_col + 1):
                x = ROWNUM_W + self._col_x[c] - self._x_off
                val = row[c] if row and c < len(row) else ""
                max_chars = max(1, (self._col_w[c] - 8) // CHAR_W)
                if len(val) > max_chars:
                    val = val[:max_chars - 1] + "…"
                cv.create_text(x + 4, y + ROW_H // 2, anchor="w", text=val,
                               font=FONT, fill="black")
            cv.create_line(0, y, W, y, fill="#e8e8e8")

        if self._sel:
            self._draw_selection_border(cv, first_col, last_col, n_vis, data_top, W, H)

        for c in range(first_col, last_col + 2):
            if c < len(self._col_x):
                x = ROWNUM_W + self._col_x[c] - self._x_off
            else:
                x = ROWNUM_W + self._total_w - self._x_off
            if x >= ROWNUM_W:
                cv.create_line(x, data_top, x, H, fill="#e8e8e8")

        self._draw_rownum_col(cv, n_vis, data_top, H)
        self._draw_header(cv, first_col, last_col, W)

        cv.create_rectangle(0, 0, ROWNUM_W, header_h, fill="#d9d9d9", outline="#a0a0a0")
        cv.create_text(ROWNUM_W // 2, header_h // 2, text="#", font=FONT, fill="#333")

        self._update_scrollbars(n_vis, W)

    def _visible_col_range(self, W):
        avail = W - ROWNUM_W
        first = self._col_at_x(ROWNUM_W)
        if first is None:
            first = 0
        last = first
        n = len(self._col_w)
        while last < n - 1:
            x = self._col_x[last + 1] - self._x_off
            if x > avail:
                break
            last += 1
        return first, min(last, n - 1)

    def _draw_rownum_col(self, cv, n_vis, data_top, H):
        cv.create_rectangle(0, data_top, ROWNUM_W, H, fill="#f3f3f3", outline="")
        for vi in range(n_vis):
            r = self._first_row + vi
            if r >= self._total_rows:
                break
            y = data_top + vi * ROW_H
            cv.create_text(ROWNUM_W - 6, y + ROW_H // 2, anchor="e",
                           text=str(r + 1), font=FONT, fill="#666")
            cv.create_line(0, y, ROWNUM_W, y, fill="#e0e0e0")
        cv.create_line(ROWNUM_W, data_top, ROWNUM_W, H, fill="#a0a0a0")

    def _padded_parts(self, col):
        """列の parts を _header_rows 段に上詰めパディング（空文字）して返す。
        最下段が個別名、上段が接頭語になるよう下詰めにする。"""
        p = self._parts[col]
        pad = self._header_rows - len(p)
        return [""] * pad + list(p)

    def _draw_header(self, cv, first_col, last_col, W):
        hr = self._header_rows
        header_h = hr * ROW_H
        cv.create_rectangle(ROWNUM_W, 0, W, header_h, fill="#d9d9d9", outline="")
        hsel0 = hsel1 = None
        if self._hsel is not None:
            hsel0, hsel1 = min(self._hsel), max(self._hsel)

        # 各列の下詰めパディング済み parts をキャッシュ
        padded = [self._padded_parts(c) for c in range(len(self._parts))]

        for row in range(hr):
            y0 = row * ROW_H
            y1 = y0 + ROW_H
            is_last = (row == hr - 1)
            if is_last:
                bg_default = "#e9e9e9"
                bg_group   = "#e9e9e9"
                text_color = "black"
            else:
                depth_ratio = row / max(hr - 1, 1)
                r_val = int(0xcf - depth_ratio * 0x20)
                b_val = int(0xe8 - depth_ratio * 0x30)
                bg_group = f"#{r_val:02x}d8{b_val:02x}"
                bg_default = "#d9d9d9"
                text_color = "#103060"

            c = first_col
            while c <= last_col:
                # row段目のテキスト（下詰め済み）
                label = padded[c][row]

                # 同じグループ（この段より上の全段が一致 かつ この段も一致）をまとめる
                key = tuple(padded[c][:row + 1])
                g_end = c
                while g_end + 1 <= last_col and tuple(padded[g_end + 1][:row + 1]) == key:
                    g_end += 1

                x0 = ROWNUM_W + self._col_x[c] - self._x_off
                x1 = ROWNUM_W + self._col_x[g_end] + self._col_w[g_end] - self._x_off
                x0c = max(x0, ROWNUM_W)
                selected = hsel0 is not None and not (g_end < hsel0 or c > hsel1)

                if selected:
                    fill = "#9cc4f4" if is_last else "#7fb0ef"
                elif label and not is_last:
                    fill = bg_group
                else:
                    fill = bg_default

                cv.create_rectangle(x0, y0, x1, y1, fill=fill, outline="#b0b0b0")
                if label:
                    cell_w = max(x1, ROWNUM_W) - x0c
                    max_chars = max(1, (cell_w - 8) // CHAR_W)
                    disp = label if len(label) <= max_chars else label[:max_chars - 1] + "…"
                    cv.create_text((x0c + x1) // 2, y0 + ROW_H // 2,
                                   text=disp, font=FONT, fill=text_color)
                c = g_end + 1

        # ヘッダ選択の枠
        if hsel0 is not None and not (hsel1 < first_col or hsel0 > last_col):
            cc0 = max(hsel0, first_col)
            cc1 = min(hsel1, last_col)
            bx0 = ROWNUM_W + self._col_x[cc0] - self._x_off
            bx1 = ROWNUM_W + self._col_x[cc1] + self._col_w[cc1] - self._x_off
            cv.create_rectangle(max(bx0, ROWNUM_W), 0, bx1, header_h,
                                outline="#1a73e8", width=2)
        cv.create_line(ROWNUM_W, header_h, W, header_h, fill="#a0a0a0")

    def _draw_crosshair(self, cv, first_col, last_col, n_vis, data_top, W, H):
        # アクティブセル（選択の末尾）の行・列に薄い十字
        _, _, ar, ac = self._sel
        # 行ハイライト
        vi = ar - self._first_row
        if 0 <= vi < n_vis:
            y = data_top + vi * ROW_H
            cv.create_rectangle(ROWNUM_W, y, W, y + ROW_H, fill="#eef5ff", outline="")
        # 列ハイライト
        if first_col <= ac <= last_col:
            x = ROWNUM_W + self._col_x[ac] - self._x_off
            cv.create_rectangle(x, data_top, x + self._col_w[ac], H, fill="#eef5ff", outline="")

    def _draw_selection(self, cv, first_col, last_col, n_vis, data_top):
        r0, c0, r1, c1 = self._norm_sel()
        for vi in range(n_vis):
            r = self._first_row + vi
            if r > r1:
                break
            if r < r0:
                continue
            y = data_top + vi * ROW_H
            for c in range(max(c0, first_col), min(c1, last_col) + 1):
                x = ROWNUM_W + self._col_x[c] - self._x_off
                cv.create_rectangle(x, y, x + self._col_w[c], y + ROW_H,
                                    fill="#cce5ff", outline="")

    def _draw_selection_border(self, cv, first_col, last_col, n_vis, data_top, W, H):
        r0, c0, r1, c1 = self._norm_sel()
        vis_top = self._first_row
        vis_bot = self._first_row + n_vis - 1
        if r1 < vis_top or r0 > vis_bot or c1 < first_col or c0 > last_col:
            return
        rr0 = max(r0, vis_top)
        rr1 = min(r1, vis_bot)
        cc0 = max(c0, first_col)
        cc1 = min(c1, last_col)
        y0 = data_top + (rr0 - self._first_row) * ROW_H
        y1 = data_top + (rr1 - self._first_row + 1) * ROW_H
        x0 = ROWNUM_W + self._col_x[cc0] - self._x_off
        x1 = ROWNUM_W + self._col_x[cc1] + self._col_w[cc1] - self._x_off
        cv.create_rectangle(max(x0, ROWNUM_W), max(y0, data_top), x1, y1,
                            outline="#1a73e8", width=2)

    def _update_scrollbars(self, n_vis, W):
        if self._total_rows:
            top = self._first_row / self._total_rows
            bot = min(1.0, (self._first_row + n_vis) / self._total_rows)
            self._vsb.set(top, bot)
        else:
            self._vsb.set(0, 1)
        avail = W - ROWNUM_W
        if self._total_w:
            left = self._x_off / self._total_w
            right = min(1.0, (self._x_off + avail) / self._total_w)
            self._hsb.set(left, right)
        else:
            self._hsb.set(0, 1)


class PlotWindow:
    """グラフ表示ポップアップ。matplotlib を Tkinter に埋め込む。

    UI フロー:
      1. 左ペインで X軸列(1つ) と Y軸列(複数) を選択
      2. 「描画」ボタン押下 → 1回だけファイルを読んでデータをキャッシュ → 描画
      3. 「Y軸を分離」チェックはキャッシュ済みデータで即再描画（再読み不要）
    """

    MAX_ROWS = 50000

    def __init__(self, parent, info):
        """info: Grid.get_plot_info() の戻り値"""
        self._info = info
        self._cached = None   # 読み込み済みデータ {col_idx: [float|nan, ...]}
        self._cached_x_col = None   # キャッシュ時のX列インデックス(None=行番号)
        self._cached_y_cols = None  # キャッシュ時のY列リスト
        self._cached_x_arr = None
        self._cached_indices = None

        try:
            from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg, NavigationToolbar2Tk
            import matplotlib.pyplot as plt
            self._plt = plt
            self._FigureCanvasTkAgg = FigureCanvasTkAgg
            self._NavToolbar = NavigationToolbar2Tk
        except ImportError:
            self._plt = None

        self._split_var = tk.BooleanVar(value=False)
        self._canvas_widget = None
        self._toolbar_widget = None

        self._win = tk.Toplevel(parent)
        self._win.title("グラフ")
        self._win.geometry("1000x660")
        self._build_ui()

    def _build_ui(self):
        if self._plt is None:
            tk.Label(self._win, text="matplotlib が必要です: pip install matplotlib",
                     fg="red").pack(padx=20, pady=20)
            return

        headers = self._info["headers"]
        sel_cols = self._info["sel_cols"]

        # ── 左ペイン: 軸選択 ──────────────────────────────────────
        left = tk.Frame(self._win, width=220, bd=1, relief=tk.GROOVE)
        left.pack(side=tk.LEFT, fill=tk.Y, padx=(4, 0), pady=4)
        left.pack_propagate(False)

        tk.Label(left, text="X軸 (1列)", font=("TkDefaultFont", 9, "bold")).pack(anchor="w", padx=6, pady=(6, 2))
        self._x_lb = tk.Listbox(left, height=6, exportselection=False, font=FONT)
        xsb = ttk.Scrollbar(left, orient=tk.VERTICAL, command=self._x_lb.yview)
        self._x_lb.config(yscrollcommand=xsb.set)
        self._x_lb.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=(6, 0))
        xsb.pack(side=tk.LEFT, fill=tk.Y, padx=(0, 4))

        tk.Label(left, text="Y軸 (複数選択可)", font=("TkDefaultFont", 9, "bold")).pack(anchor="w", padx=6, pady=(8, 2))
        self._y_lb = tk.Listbox(left, selectmode=tk.EXTENDED, exportselection=False, font=FONT)
        ysb = ttk.Scrollbar(left, orient=tk.VERTICAL, command=self._y_lb.yview)
        self._y_lb.config(yscrollcommand=ysb.set)
        self._y_lb.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=(6, 0))
        ysb.pack(side=tk.LEFT, fill=tk.Y, padx=(0, 4))

        # リストに全列名を投入
        self._x_lb.insert(tk.END, "(行番号)")
        for h in headers:
            self._x_lb.insert(tk.END, h)
        self._x_lb.selection_set(0)  # デフォルト: 行番号

        for h in headers:
            self._y_lb.insert(tk.END, h)
        # 選択列を初期選択
        for c in sel_cols:
            self._y_lb.selection_set(c)
        if sel_cols:
            self._y_lb.see(sel_cols[0])

        # ── 下部コントロール ──────────────────────────────────────
        btn_frame = tk.Frame(left)
        btn_frame.pack(side=tk.BOTTOM, fill=tk.X, padx=6, pady=6)

        tk.Checkbutton(btn_frame, text="Y軸を分離", variable=self._split_var,
                       command=self._on_split_changed).pack(anchor="w")

        self._status_lbl = tk.Label(btn_frame, text="", fg="#666", font=("TkDefaultFont", 8),
                                    wraplength=200, justify=tk.LEFT)
        self._status_lbl.pack(anchor="w", pady=(2, 4))

        tk.Button(btn_frame, text="描画", font=("TkDefaultFont", 10, "bold"),
                  command=self._on_draw).pack(fill=tk.X)

        # ── 右ペイン: グラフ ──────────────────────────────────────
        self._fig_frame = tk.Frame(self._win)
        self._fig_frame.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=4, pady=4)

    def _get_selections(self):
        """(x_col_or_None, [y_col_indices]) を返す。x_col_or_None=None は行番号。"""
        xi = self._x_lb.curselection()
        x_col = None if (not xi or xi[0] == 0) else xi[0] - 1  # 0番目=(行番号)

        yi = self._y_lb.curselection()
        y_cols = list(yi)  # Listbox は headers と同インデックス
        return x_col, y_cols

    def _on_draw(self):
        x_col, y_cols = self._get_selections()
        if not y_cols:
            self._status_lbl.config(text="Y軸列を1つ以上選択してください")
            return

        # 必要な列セット
        need_cols = set(y_cols)
        if x_col is not None:
            need_cols.add(x_col)

        # キャッシュ有効チェック: 必要な列が全てキャッシュ済みか
        if (self._cached is not None
                and need_cols <= set(self._cached.keys())
                and self._cached_indices is not None):
            self._status_lbl.config(text="キャッシュ済みデータで描画")
            self._draw(x_col, y_cols)
            return

        # ── ファイル読み込み ──
        total = self._info["total_rows"]
        sampled = total > self.MAX_ROWS
        if sampled:
            step = total / self.MAX_ROWS
            indices = [int(i * step) for i in range(self.MAX_ROWS)]
        else:
            indices = list(range(total))

        self._status_lbl.config(text="読み込み中…")
        self._win.update_idletasks()

        import csv as _csv
        cols = {c: [] for c in need_cols}
        try:
            with open(self._info["filepath"], "r",
                      encoding=self._info["encoding"], newline="") as f:
                offsets = self._info["offsets"]
                for idx in indices:
                    f.seek(offsets[idx])
                    line = f.readline()
                    fields = next(_csv.reader([line])) if line else []
                    for c in need_cols:
                        raw = fields[c] if c < len(fields) else ""
                        try:
                            cols[c].append(float(raw))
                        except (ValueError, OverflowError):
                            cols[c].append(float("nan"))
        except Exception as e:
            self._status_lbl.config(text=f"読み込みエラー: {e}")
            return

        self._cached = cols
        self._cached_indices = indices
        n_pts = len(indices)
        msg = f"{n_pts:,} pts"
        if sampled:
            msg += f"\n(全{total:,}行からサンプリング)"
        self._status_lbl.config(text=msg)
        self._draw(x_col, y_cols)

    def _on_split_changed(self):
        # キャッシュがあれば即再描画、なければ何もしない
        if self._cached is not None and self._cached_y_cols is not None:
            self._draw(self._cached_x_col, self._cached_y_cols)

    def _draw(self, x_col, y_cols):
        self._cached_x_col = x_col
        self._cached_y_cols = y_cols

        # 既存ウィジェット破棄
        if self._canvas_widget:
            self._canvas_widget.get_tk_widget().destroy()
        if self._toolbar_widget:
            self._toolbar_widget.destroy()

        headers = self._info["headers"]
        indices = self._cached_indices
        split = self._split_var.get()
        n = len(y_cols)

        # X軸データ
        if x_col is None:
            x_arr = [i + 1 for i in indices]
            x_label = "行番号"
        else:
            x_arr = self._cached[x_col]
            x_label = headers[x_col]

        fig_h = max(4.0, 1.6 * n) if split else 4.5
        fig, axes = self._plt.subplots(
            n if split else 1, 1,
            figsize=(9, fig_h),
            sharex=True,
            squeeze=False,
        )
        fig.subplots_adjust(
            hspace=0.06 if split else 0.25,
            left=0.10, right=0.97,
            top=0.94, bottom=0.07,
        )
        colors = self._plt.rcParams["axes.prop_cycle"].by_key()["color"]

        for i, col in enumerate(y_cols):
            ax = axes[i][0] if split else axes[0][0]
            y_arr = self._cached[col]
            label = headers[col]
            color = colors[i % len(colors)]
            ax.plot(x_arr, y_arr, linewidth=0.8, color=color,
                    label=label if not split else None)
            if split:
                ax.set_ylabel(label, fontsize=8, labelpad=2)
                ax.tick_params(labelsize=7)
                ax.grid(True, linewidth=0.4, alpha=0.5)

        if not split:
            ax = axes[0][0]
            if n <= 8:
                ax.legend(fontsize=7, loc="best")
            ax.set_xlabel(x_label, fontsize=8)
            ax.grid(True, linewidth=0.4, alpha=0.5)
            ax.tick_params(labelsize=7)
        else:
            axes[-1][0].set_xlabel(x_label, fontsize=8)
            axes[-1][0].tick_params(labelsize=7)

        canvas = self._FigureCanvasTkAgg(fig, master=self._fig_frame)
        canvas.draw()
        canvas.get_tk_widget().pack(fill=tk.BOTH, expand=True)
        toolbar = self._NavToolbar(canvas, self._fig_frame)
        toolbar.update()
        self._canvas_widget = canvas
        self._toolbar_widget = toolbar
        self._plt.close(fig)


class App:
    """ツールバー + 複数タブ(Notebook)を管理するアプリ本体。"""

    def __init__(self, root, filepaths=None):
        self.root = root
        root.title("CSV Viewer")
        root.geometry("1200x700")
        self._grids = []  # 開いている Grid のリスト

        self._build_ui()
        self._setup_dnd()

        if filepaths:
            root.after(100, lambda: [self.open_file_in_new_tab(p) for p in filepaths])
        else:
            self._new_tab()  # 空のタブ

    # ── UI ────────────────────────────────────────────────────────
    def _build_ui(self):
        toolbar = tk.Frame(self.root, bd=1, relief=tk.RAISED)
        toolbar.pack(side=tk.TOP, fill=tk.X)
        tk.Button(toolbar, text="Open CSV", command=self._browse).pack(side=tk.LEFT, padx=4, pady=2)
        tk.Button(toolbar, text="Copy", command=self._copy).pack(side=tk.LEFT, padx=2, pady=2)
        tk.Button(toolbar, text="Copy+ラベル", command=self._copy_labels).pack(side=tk.LEFT, padx=2, pady=2)
        tk.Button(toolbar, text="タブを閉じる", command=self._close_tab).pack(side=tk.LEFT, padx=2, pady=2)
        tk.Button(toolbar, text="グラフ", command=self._open_plot).pack(side=tk.LEFT, padx=(8, 2), pady=2)

        tk.Label(toolbar, text="検索:").pack(side=tk.LEFT, padx=(12, 2))
        self._search_var = tk.StringVar()
        ent = tk.Entry(toolbar, textvariable=self._search_var, width=20)
        ent.pack(side=tk.LEFT, padx=2)
        ent.bind("<Return>", lambda e: self._do_search(1))
        ent.bind("<Shift-Return>", lambda e: self._do_search(-1))
        tk.Button(toolbar, text="▲", width=2, command=lambda: self._do_search(-1)).pack(side=tk.LEFT)
        tk.Button(toolbar, text="▼", width=2, command=lambda: self._do_search(1)).pack(side=tk.LEFT, padx=(0, 6))

        self._status = tk.StringVar(value="No file loaded")
        tk.Label(toolbar, textvariable=self._status, anchor=tk.W).pack(side=tk.LEFT, padx=8)

        self._nb = ttk.Notebook(self.root)
        self._nb.pack(fill=tk.BOTH, expand=True)
        self._nb.bind("<<NotebookTabChanged>>", self._on_tab_changed)
        self._nb.bind("<Button-2>", self._on_tab_middle_click)

        self.root.bind("<Control-w>", lambda e: self._close_tab())
        self.root.bind("<Control-t>", lambda e: self._new_tab())

    def _setup_dnd(self):
        # Notebook 全体でもドロップ受付（空タブ時など）
        try:
            from tkinterdnd2 import DND_FILES
            self._nb.drop_target_register(DND_FILES)
            self._nb.dnd_bind("<<Drop>>", self._on_drop)
        except Exception:
            pass

    def _on_drop(self, event):
        for path in parse_drop(event.data):
            self.open_file_in_new_tab(path)

    # ── タブ管理 ─────────────────────────────────────────────────
    def _new_tab(self, title="(empty)"):
        grid = Grid(self._nb, set_status=self._set_status,
                    open_in_new_tab=self.open_file_in_new_tab)
        self._grids.append(grid)
        self._nb.add(grid, text=title)
        self._nb.select(grid)
        grid.after(50, grid.focus_grid)
        return grid

    def open_file_in_new_tab(self, path):
        cur = self._current_grid()
        # 現在のタブが空なら再利用、そうでなければ新規タブ
        if cur is not None and cur._filepath is None:
            grid = cur
        else:
            grid = self._new_tab()
        if grid.open_file(path):
            self._nb.tab(grid, text=os.path.basename(path))
            self._nb.select(grid)
            grid.after(50, grid.focus_grid)

    def _close_tab(self):
        grid = self._current_grid()
        if grid is None:
            return
        self._nb.forget(grid)
        if grid in self._grids:
            self._grids.remove(grid)
        grid.destroy()
        if not self._nb.tabs():
            self._new_tab()

    def _on_tab_middle_click(self, event):
        try:
            idx = self._nb.index("@%d,%d" % (event.x, event.y))
        except Exception:
            return
        tab_id = self._nb.tabs()[idx]
        widget = self.root.nametowidget(tab_id)
        self._nb.forget(widget)
        if widget in self._grids:
            self._grids.remove(widget)
        widget.destroy()
        if not self._nb.tabs():
            self._new_tab()

    def _current_grid(self):
        cur = self._nb.select()
        if not cur:
            return None
        return self.root.nametowidget(cur)

    def _on_tab_changed(self, event):
        grid = self._current_grid()
        if grid is not None:
            if grid._filepath:
                self._set_status(grid, f"{grid._filepath}  "
                                 f"{grid._total_rows} rows × {len(grid._headers)} cols")
            else:
                self._status.set("No file loaded")
            grid.after(10, grid.focus_grid)

    # ── ツールバー操作の委譲 ─────────────────────────────────────
    def _browse(self):
        paths = filedialog.askopenfilenames(
            filetypes=[("CSV files", "*.csv"), ("All files", "*.*")])
        for p in paths:
            self.open_file_in_new_tab(p)

    def _copy(self):
        g = self._current_grid()
        if g:
            g.copy_selection()

    def _copy_labels(self):
        g = self._current_grid()
        if g:
            g.copy_with_labels()

    def _open_plot(self):
        g = self._current_grid()
        if not g:
            return
        info = g.get_plot_info()
        if info is None:
            messagebox.showinfo("グラフ", "列を選択してからグラフボタンを押してください。")
            return
        PlotWindow(self.root, info)

    def _do_search(self, direction):
        g = self._current_grid()
        if g:
            g.search(self._search_var.get(), direction)
        return "break"

    def _set_status(self, grid, text):
        # 現在表示中のタブのときのみ反映
        if grid is self._current_grid():
            self._status.set(text)


def main():
    paths = [to_local_path(p) for p in sys.argv[1:]] or None
    root = _make_tk()
    App(root, paths)
    root.mainloop()


if __name__ == "__main__":
    main()

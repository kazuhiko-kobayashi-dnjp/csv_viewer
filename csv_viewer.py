"""軽量CSVビューワ (Canvas仮想スクロール版)

特徴:
- 行番号列 + 列ヘッダ表示
- Excelライクな矩形セル選択（ドラッグ）
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
HEADER_ROWS = 2       # 上段:接頭語 下段:個別名
FONT = ("TkDefaultFont", 9)
CACHE_LIMIT = 4000    # 行キャッシュ上限


def _make_tk():
    try:
        from tkinterdnd2 import TkinterDnD
        return TkinterDnD.Tk()
    except Exception:
        return tk.Tk()


def split_label(name):
    """ラベルを (接頭語, 個別名) に分割。最初のドットで区切る。"""
    if "." in name:
        prefix, indiv = name.split(".", 1)
        return prefix, indiv
    return "", name


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


class CSVViewer:
    def __init__(self, root, filepath=None):
        self.root = root
        root.title("CSV Viewer")
        root.geometry("1200x700")

        self._filepath = None
        self._encoding = "utf-8"
        self._headers = []          # 列名リスト
        self._prefixes = []         # 各列の接頭語
        self._indivs = []           # 各列の個別名
        self._col_w = []            # 各列の幅
        self._col_x = []            # 各列の左端x(データ座標, 累積)
        self._total_w = 0           # データ列の総幅

        self._offsets = [0]         # 各データ行の開始バイトオフセット (索引)
        self._total_rows = 0        # 索引済みデータ行数
        self._indexing = False

        self._row_cache = {}        # row_idx -> list[str]
        self._cache_order = []

        self._first_row = 0         # 表示先頭データ行
        self._x_off = 0             # 水平スクロール(px, データ座標)

        # 選択範囲 (データ座標: r0,c0 - r1,c1)
        self._sel = None
        self._dragging = False

        self._build_ui()
        self._setup_dnd()
        if filepath:
            root.after(100, lambda: self._open_file(filepath))

    # ── UI ────────────────────────────────────────────────────────
    def _build_ui(self):
        toolbar = tk.Frame(self.root, bd=1, relief=tk.RAISED)
        toolbar.pack(side=tk.TOP, fill=tk.X)
        tk.Button(toolbar, text="Open CSV", command=self._browse).pack(side=tk.LEFT, padx=4, pady=2)
        tk.Button(toolbar, text="Copy", command=self._copy_selection).pack(side=tk.LEFT, padx=4, pady=2)

        tk.Label(toolbar, text="検索:").pack(side=tk.LEFT, padx=(12, 2))
        self._search_var = tk.StringVar()
        ent = tk.Entry(toolbar, textvariable=self._search_var, width=20)
        ent.pack(side=tk.LEFT, padx=2)
        ent.bind("<Return>", lambda e: self._search(1))
        ent.bind("<Shift-Return>", lambda e: self._search(-1))
        tk.Button(toolbar, text="\u25b2", width=2,
                  command=lambda: self._search(-1)).pack(side=tk.LEFT)
        tk.Button(toolbar, text="\u25bc", width=2,
                  command=lambda: self._search(1)).pack(side=tk.LEFT, padx=(0, 6))

        self._status = tk.StringVar(value="No file loaded")
        tk.Label(toolbar, textvariable=self._status, anchor=tk.W).pack(side=tk.LEFT, padx=8)

        container = tk.Frame(self.root)
        container.pack(fill=tk.BOTH, expand=True)

        self._canvas = tk.Canvas(container, bg="white", highlightthickness=0)
        self._vsb = ttk.Scrollbar(container, orient=tk.VERTICAL, command=self._on_vscroll)
        self._hsb = ttk.Scrollbar(container, orient=tk.HORIZONTAL, command=self._on_hscroll)
        self._canvas.grid(row=0, column=0, sticky="nsew")
        self._vsb.grid(row=0, column=1, sticky="ns")
        self._hsb.grid(row=1, column=0, sticky="ew")
        container.rowconfigure(0, weight=1)
        container.columnconfigure(0, weight=1)

        self._canvas.bind("<Configure>", lambda e: self._redraw())
        self._canvas.bind("<MouseWheel>", self._on_wheel)
        self._canvas.bind("<Button-4>", self._on_wheel)
        self._canvas.bind("<Button-5>", self._on_wheel)
        self._canvas.bind("<Shift-MouseWheel>", self._on_wheel_h)
        self._canvas.bind("<Button-1>", self._on_press)
        self._canvas.bind("<B1-Motion>", self._on_drag)
        self._canvas.bind("<ButtonRelease-1>", self._on_release)
        self.root.bind("<Control-c>", lambda e: self._copy_selection())
        for key, dr, dc in (("Up", -1, 0), ("Down", 1, 0),
                            ("Left", 0, -1), ("Right", 0, 1)):
            self._canvas.bind(f"<{key}>",
                              lambda e, a=dr, b=dc: self._move_sel(a, b))
        self._canvas.bind("<Prior>", lambda e: self._page(-1))
        self._canvas.bind("<Next>", lambda e: self._page(1))

    def _setup_dnd(self):
        try:
            from tkinterdnd2 import DND_FILES
            self._canvas.drop_target_register(DND_FILES)
            self._canvas.dnd_bind("<<Drop>>", self._on_drop)
        except Exception:
            pass

    def _on_drop(self, event):
        raw = event.data
        # 複数ファイルがブレースで来る場合は先頭のみ
        if raw.startswith("{"):
            raw = raw[1:].split("}", 1)[0]
        else:
            raw = raw.split(" ")[0]
        path = to_local_path(raw)
        if path:
            self._open_file(path)

    def _browse(self):
        path = filedialog.askopenfilename(
            filetypes=[("CSV files", "*.csv"), ("All files", "*.*")])
        if path:
            self._open_file(path)

    # ── ファイル読込 ─────────────────────────────────────────────
    def _open_file(self, path):
        try:
            enc = self._detect_encoding(path)
            with open(path, "r", encoding=enc, newline="") as f:
                header_line = f.readline()
            if not header_line:
                messagebox.showerror("Error", "空のファイルです")
                return
            headers = next(csv.reader([header_line]))
        except Exception as e:
            messagebox.showerror("Error", f"読込失敗: {e}")
            return

        self._filepath = path
        self._encoding = enc
        self._headers = headers
        self._prefixes = []
        self._indivs = []
        for h in headers:
            p, i = split_label(h)
            self._prefixes.append(p)
            self._indivs.append(i)
        self._compute_columns()

        self._row_cache.clear()
        self._cache_order.clear()
        self._offsets = []
        self._total_rows = 0
        self._first_row = 0
        self._x_off = 0
        self._sel = None

        self._status.set(f"{path}  ({len(headers)} cols)  索引中…")
        self._start_indexing(path, enc, len(header_line.encode(enc)))

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
        for i in range(len(self._headers)):
            label_len = max(len(self._prefixes[i]), len(self._indivs[i]))
            w = max(COL_MIN_W, min(COL_MAX_W, label_len * CHAR_W + 12))
            self._col_w.append(w)
            self._col_x.append(x)
            x += w
        self._total_w = x

    def _start_indexing(self, path, enc, header_bytes):
        self._indexing = True
        self._pending = None          # worker -> main 受け渡し
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
        self.root.after(100, self._poll_index)

    def _poll_index(self):
        if self._index_error is not None:
            self._status.set(f"\u7d22\u5f15\u30a8\u30e9\u30fc: {self._index_error}")
            self._indexing = False
            return
        if self._pending is not None:
            offsets, done = self._pending
            self._pending = None
            self._offsets = offsets
            self._total_rows = len(offsets)
            if done:
                self._indexing = False
                self._status.set(
                    f"{self._filepath}  {self._total_rows} rows \u00d7 {len(self._headers)} cols")
                self._redraw()
                return
            self._status.set(
                f"{self._filepath}  \u7d22\u5f15\u4e2d\u2026 {self._total_rows} rows")
            self._redraw()
        if self._indexing:
            self.root.after(100, self._poll_index)

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
        return max(1, (h - HEADER_ROWS * ROW_H) // ROW_H)

    def _on_vscroll(self, *args):
        if not self._total_rows:
            return
        if args[0] == "moveto":
            frac = float(args[1])
            self._first_row = int(frac * self._total_rows)
        elif args[0] == "scroll":
            n = int(args[1])
            if args[2] == "pages":
                n *= self._visible_data_rows()
            self._first_row += n
        self._clamp_scroll()
        self._redraw()

    def _on_hscroll(self, *args):
        if args[0] == "moveto":
            frac = float(args[1])
            self._x_off = int(frac * self._total_w)
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
        delta = -int(event.delta / 120) * 40
        self._x_off += delta
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
        """キャンバスx(px) -> データ列index または None。"""
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
        top = HEADER_ROWS * ROW_H
        if py < top:
            return None
        r = self._first_row + (py - top) // ROW_H
        if r >= self._total_rows:
            return None
        return int(r)

    # ── 選択 ─────────────────────────────────────────────────────
    def _on_press(self, event):
        c = self._col_at_x(event.x)
        r = self._row_at_y(event.y)
        if c is None or r is None:
            return
        self._sel = (r, c, r, c)
        self._dragging = True
        self._canvas.focus_set()
        self._redraw()

    def _on_drag(self, event):
        if not self._dragging or self._sel is None:
            return
        if event.y > self._canvas.winfo_height() - ROW_H:
            self._first_row += 1
        elif event.y < HEADER_ROWS * ROW_H:
            self._first_row -= 1
        self._clamp_scroll()
        cx = max(ROWNUM_W, min(event.x, self._canvas.winfo_width() - 1))
        cy = max(HEADER_ROWS * ROW_H, min(event.y, self._canvas.winfo_height() - 1))
        c = self._col_at_x(cx)
        r = self._row_at_y(cy)
        if c is None:
            c = self._sel[3]
        if r is None:
            r = self._sel[2]
        self._sel = (self._sel[0], self._sel[1], r, c)
        self._redraw()

    def _on_release(self, event):
        self._dragging = False

    def _norm_sel(self):
        r0, c0, r1, c1 = self._sel
        return min(r0, r1), min(c0, c1), max(r0, r1), max(c0, c1)

    def _copy_selection(self):
        if not self._sel:
            return
        r0, c0, r1, c1 = self._norm_sel()
        lines = []
        for r in range(r0, r1 + 1):
            row = self._get_row(r) or []
            cells = [row[c] if c < len(row) else "" for c in range(c0, c1 + 1)]
            lines.append("\t".join(cells))
        text = "\n".join(lines)
        self.root.clipboard_clear()
        self.root.clipboard_append(text)

    # ── キーボード移動 ───────────────────────────────────────────
    def _move_sel(self, dr, dc):
        if not self._headers or not self._total_rows:
            return "break"
        if self._sel is None:
            r, c = self._first_row, 0
        else:
            _, _, r, c = self._sel
        r = max(0, min(r + dr, self._total_rows - 1))
        c = max(0, min(c + dc, len(self._headers) - 1))
        self._sel = (r, c, r, c)
        self._ensure_visible(r, c)
        self._redraw()
        return "break"

    def _page(self, direction):
        n = self._visible_data_rows()
        self._first_row += direction * n
        self._clamp_scroll()
        self._redraw()
        return "break"

    def _ensure_visible(self, r, c):
        n_vis = self._visible_data_rows()
        if r < self._first_row:
            self._first_row = r
        elif r >= self._first_row + n_vis:
            self._first_row = r - n_vis + 1
        # 水平
        x0 = self._col_x[c]
        x1 = x0 + self._col_w[c]
        avail = self._canvas.winfo_width() - ROWNUM_W
        if x0 < self._x_off:
            self._x_off = x0
        elif x1 > self._x_off + avail:
            self._x_off = x1 - avail
        self._clamp_scroll()

    # ── 検索 ─────────────────────────────────────────────────────
    def _search(self, direction):
        term = self._search_var.get().strip()
        if not term or not self._total_rows:
            return "break"
        # 空白区切りの複数キーワードをすべて含めばヒット(AND、部分一致、大小無視)
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
        # 列名(ラベル)も検索対象に含める
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
                # 各列の検索対象テキスト = セル値 + 列ラベル
                def text_at(cc):
                    cell = lowered[cc] if cc < len(lowered) else ""
                    label = header_lower[cc] if cc < len(header_lower) else ""
                    return cell + " " + label
                if single:
                    # 同一行内をセル単位で走査
                    cols = range(c + 1, ncol) if direction > 0 else range(c - 1, -1, -1)
                    for cc in cols:
                        if terms[0] in text_at(cc):
                            self._search_result = (r, cc)
                            return
                else:
                    # 複数キーワードは「行内のどこかに全部あれば」ヒット
                    texts = [text_at(cc) for cc in range(ncol)]
                    if all(any(t in tx for tx in texts) for t in terms):
                        for cc in range(ncol):
                            if any(t in texts[cc] for t in terms):
                                self._search_result = (r, cc)
                                return
                # 次の行へ
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
        self._status.set(f"\u691c\u7d22\u4e2d: {term}")
        threading.Thread(target=thread, daemon=True).start()
        self.root.after(50, self._poll_search)
        return "break"

    def _poll_search(self):
        res = getattr(self, "_search_result", None)
        if res is None:
            self.root.after(50, self._poll_search)
            return
        if res == "notfound":
            self._status.set(f"\u300c{self._search_term}\u300d\u306f\u898b\u3064\u304b\u308a\u307e\u305b\u3093")
            return
        self._goto_match(res[0], res[1])


    def _goto_match(self, r, c):
        self._sel = (r, c, r, c)
        self._ensure_visible(r, c)
        self._status.set(f"\u4e00\u81f4: \u884c{r + 1} \u5217{c + 1} ({self._headers[c]})")
        self._canvas.focus_set()
        self._redraw()

    # ── 描画 ─────────────────────────────────────────────────────
    def _redraw(self):
        cv = self._canvas
        cv.delete("all")
        if not self._headers:
            cv.create_text(20, 20, anchor="nw", text="CSVファイルを開いてください",
                           fill="gray", font=FONT)
            return

        W = cv.winfo_width()
        H = cv.winfo_height()
        n_vis = self._visible_data_rows()
        header_h = HEADER_ROWS * ROW_H
        data_top = header_h

        first_col, last_col = self._visible_col_range(W)

        if self._sel:
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

    def _draw_header(self, cv, first_col, last_col, W):
        cv.create_rectangle(ROWNUM_W, 0, W, HEADER_ROWS * ROW_H,
                            fill="#d9d9d9", outline="")
        # 下段: 個別名
        for c in range(first_col, last_col + 1):
            x0 = ROWNUM_W + self._col_x[c] - self._x_off
            x1 = x0 + self._col_w[c]
            cv.create_rectangle(x0, ROW_H, x1, HEADER_ROWS * ROW_H,
                                fill="#e9e9e9", outline="#b0b0b0")
            cv.create_text((x0 + x1) // 2, ROW_H + ROW_H // 2,
                           text=self._indivs[c], font=FONT, fill="black")
        # 上段: 接頭語を連続する同一接頭語でまとめる
        c = first_col
        while c <= last_col:
            p = self._prefixes[c]
            g_end = c
            while g_end + 1 <= last_col and self._prefixes[g_end + 1] == p:
                g_end += 1
            x0 = ROWNUM_W + self._col_x[c] - self._x_off
            x1 = ROWNUM_W + self._col_x[g_end] + self._col_w[g_end] - self._x_off
            x0c = max(x0, ROWNUM_W)
            cv.create_rectangle(x0, 0, x1, ROW_H,
                                fill="#cfd8e8" if p else "#d9d9d9",
                                outline="#b0b0b0")
            if p:
                cv.create_text((x0c + x1) // 2, ROW_H // 2, text=p,
                               font=FONT, fill="#103060")
            c = g_end + 1
        cv.create_line(ROWNUM_W, HEADER_ROWS * ROW_H, W, HEADER_ROWS * ROW_H,
                       fill="#a0a0a0")

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


def main():
    path = sys.argv[1] if len(sys.argv) > 1 else None
    root = _make_tk()
    CSVViewer(root, path)
    root.mainloop()


if __name__ == "__main__":
    main()

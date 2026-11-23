"""
Monthly workbook consolidator
=============================

Folder layout handled::

    root/
      <month>/                (e.g. 2025-01, Jan-2025 ...)
        <data type>/          (e.g. Sales, Inventory ...)
          file1.xlsx, file2.xlsx, ...   (many sheets each, 1..n tables per sheet)

For every month a consolidated workbook is produced.  Inside it, every sheet
name found in the files of a data-type sub-folder is consolidated across all
those files (sheet A of file1 + sheet A of file2 + ...), then sheet B, etc.

Table detection (per sheet, no configuration needed)
----------------------------------------------------
1. Recursive X-Y cut: the used grid is split on blank rows / blank columns
   until only contiguous rectangles remain -> handles stacked and
   side-by-side tables.
2. Fragments that continue a table above them (blank line inside a table,
   repeated header, detached totals line) are glued back on.
3. Inside each rectangle the header row is located (title / preamble rows
   above it are ignored, merged multi-level headers are flattened with
   " | "), then totals, repeated headers and trailing footnotes are dropped.

Cross-file matching
-------------------
Tables are matched to "templates" per (data type, sheet name) by header
similarity, so column order may differ between files, columns may be added,
and a sheet holding two tables is consolidated into two outputs.
"""
from __future__ import annotations

import csv
import datetime as dt
import math
import re
from collections import OrderedDict
from dataclasses import dataclass, field, fields
from decimal import Decimal
from pathlib import Path
from typing import Callable, Optional

from openpyxl import Workbook, load_workbook
from openpyxl.cell import WriteOnlyCell
from openpyxl.cell.cell import ILLEGAL_CHARACTERS_RE
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

EXCEL_ROW_LIMIT = 1_000_000
SUPPORTED_EXT = (".xlsx", ".xlsm", ".xls", ".csv")
TOTAL_RE = re.compile(r"^\s*(grand\s+|sub[\s\-]*)?totals?\b", re.I)


# --------------------------------------------------------------------------
# options
# --------------------------------------------------------------------------
@dataclass
class Options:
    row_gap: int = 1              # blank rows that separate two tables
    col_gap: int = 1              # blank columns that separate two tables
    min_data_rows: int = 1        # tables with fewer data rows are treated as notes
    match_threshold: float = 0.5  # header similarity needed to be "the same table"
    skip_totals: bool = True      # drop "Total / Grand total / Subtotal" rows
    skip_hidden_sheets: bool = False
    add_source_file: bool = True
    add_source_row: bool = False
    multi_table_layout: str = "separate"   # "separate" sheets or "stacked" in one sheet
    per_type_files: bool = False  # one workbook per month+type instead of per month
    extensions: tuple = SUPPORTED_EXT

    @classmethod
    def from_dict(cls, d: Optional[dict]) -> "Options":
        d = d or {}
        o = cls()
        for f in fields(cls):
            if f.name not in d or d[f.name] is None:
                continue
            v = d[f.name]
            cur = getattr(o, f.name)
            try:
                if isinstance(cur, bool):
                    v = v if isinstance(v, bool) else str(v).lower() in ("1", "true", "yes", "on")
                elif isinstance(cur, int):
                    v = int(v)
                elif isinstance(cur, float):
                    v = float(v)
                elif f.name == "extensions":
                    v = tuple(x if x.startswith(".") else "." + x for x in v)
                else:
                    v = str(v)
            except (TypeError, ValueError):
                continue
            setattr(o, f.name, v)
        o.row_gap = max(1, o.row_gap)
        o.col_gap = max(1, o.col_gap)
        o.min_data_rows = max(1, o.min_data_rows)
        return o


class Cancelled(Exception):
    pass


# --------------------------------------------------------------------------
# reading
# --------------------------------------------------------------------------
@dataclass
class SheetData:
    name: str
    grid: list
    merged: set = field(default_factory=set)   # cells filled from a merged range (not top-left)
    hidden: bool = False


def _clean(v):
    if v is None:
        return None
    if isinstance(v, str):
        v = ILLEGAL_CHARACTERS_RE.sub("", v).replace("\xa0", " ").strip()
        return v or None
    if isinstance(v, float) and v != v:
        return None
    return v


def _pad(rows):
    w = max((len(r) for r in rows), default=0)
    for r in rows:
        if len(r) < w:
            r.extend([None] * (w - len(r)))
    while rows and not any(v is not None for v in rows[-1]):
        rows.pop()
    return rows


def _read_openpyxl(path: Path, opts: Options):
    big = path.stat().st_size > 40 * 1024 * 1024
    wb = load_workbook(path, data_only=True, read_only=big)
    sheets = []
    try:
        for ws in wb.worksheets:
            rows = []
            for r in ws.iter_rows(values_only=True):
                r = [_clean(v) for v in r]
                while r and r[-1] is None:
                    r.pop()
                rows.append(r)
            grid = _pad(rows)
            merged = set()
            if not big:
                for rng in ws.merged_cells.ranges:
                    n = (rng.max_row - rng.min_row + 1) * (rng.max_col - rng.min_col + 1)
                    if n > 20000 or rng.min_row > len(grid) or not grid:
                        continue
                    tl = grid[rng.min_row - 1][rng.min_col - 1] if rng.min_col <= len(grid[0]) else None
                    for rr in range(rng.min_row, min(rng.max_row, len(grid)) + 1):
                        for cc in range(rng.min_col, min(rng.max_col, len(grid[0])) + 1):
                            if (rr, cc) != (rng.min_row, rng.min_col):
                                if grid[rr - 1][cc - 1] is None:
                                    grid[rr - 1][cc - 1] = tl
                                merged.add((rr - 1, cc - 1))
            sheets.append(SheetData(ws.title, grid, merged, ws.sheet_state != "visible"))
    finally:
        wb.close()
    return sheets


def _read_xls(path: Path, opts: Options):
    import pandas as pd
    frames = pd.read_excel(path, sheet_name=None, header=None, engine="xlrd")
    out = []
    for name, df in frames.items():
        df = df.astype(object).where(df.notna(), None)
        grid = _pad([[_clean(v) for v in row] for row in df.values.tolist()])
        out.append(SheetData(str(name), grid))
    return out


def _read_csv(path: Path, opts: Options):
    raw = path.read_bytes()
    for enc in ("utf-8-sig", "cp1252", "latin-1"):
        try:
            text = raw.decode(enc)
            break
        except UnicodeDecodeError:
            continue
    try:
        dialect = csv.Sniffer().sniff(text[:20000], delimiters=",;\t|")
    except csv.Error:
        dialect = csv.excel

    def conv(s):
        s = _clean(s)
        if s is None:
            return None
        if re.fullmatch(r"-?(0|[1-9]\d*)", s):
            return int(s)
        if re.fullmatch(r"-?\d+\.\d+", s):
            return float(s)
        return s

    rows = [[conv(c) for c in r] for r in csv.reader(text.splitlines(), dialect)]
    return [SheetData("csv", _pad(rows))]


def read_sheets(path: Path, opts: Options) -> list:
    ext = path.suffix.lower()
    if ext in (".xlsx", ".xlsm"):
        return _read_openpyxl(path, opts)
    if ext == ".xls":
        return _read_xls(path, opts)
    if ext == ".csv":
        return _read_csv(path, opts)
    raise ValueError(f"unsupported file type {ext}")


# --------------------------------------------------------------------------
# table detection
# --------------------------------------------------------------------------
def norm(s) -> str:
    s = str(s).lower()
    n = re.sub(r"[\W_]+", "", s)
    return n or s.strip()


def _hdr_text(v) -> str:
    if isinstance(v, dt.datetime):
        return v.strftime("%Y-%m-%d") if v.time() == dt.time(0) else v.isoformat(sep=" ")
    if isinstance(v, dt.date):
        return v.isoformat()
    if isinstance(v, float) and v.is_integer():
        return str(int(v))
    return re.sub(r"\s+", " ", str(v)).strip()


def _is_num_or_date(v) -> bool:
    return (isinstance(v, (int, float, Decimal)) and not isinstance(v, bool)) or isinstance(v, (dt.date, dt.datetime, dt.time))


@dataclass
class Table:
    header: list
    norms: list
    rows: list
    rownums: list
    r0: int
    r1: int
    c0: int
    c1: int
    title: Optional[str] = None
    header_row: int = 0

    @property
    def ref(self) -> str:
        return f"{get_column_letter(self.c0 + 1)}{self.r0 + 1}:{get_column_letter(self.c1)}{self.r1}"


def _split_axis(flags, gap):
    spans, start, last, blanks = [], None, None, 0
    for i, f in enumerate(flags):
        if f:
            if start is None:
                start = i
            elif blanks >= gap:
                spans.append((start, last + 1))
                start = i
            blanks, last = 0, i
        else:
            blanks += 1
    if start is not None:
        spans.append((start, last + 1))
    return spans


def find_regions(grid, row_gap=1, col_gap=1):
    """Recursive X-Y cut -> list of (r0, r1, c0, c1) half-open rectangles."""
    if not grid:
        return []
    out = []
    stack = [(0, len(grid), 0, len(grid[0]))]
    while stack:
        r0, r1, c0, c1 = stack.pop()
        rflags = [any(grid[r][c] is not None for c in range(c0, c1)) for r in range(r0, r1)]
        rs = _split_axis(rflags, row_gap)
        if not rs:
            continue
        if len(rs) > 1 or rs[0] != (0, r1 - r0):
            stack.extend((r0 + s, r0 + e, c0, c1) for s, e in rs)
            continue
        cflags = [any(grid[r][c] is not None for r in range(r0, r1)) for c in range(c0, c1)]
        cs = _split_axis(cflags, col_gap)
        if len(cs) > 1 or cs[0] != (0, c1 - c0):
            stack.extend((r0, r1, c0 + s, c0 + e) for s, e in cs)
            continue
        out.append((r0, r1, c0, c1))
    out.sort(key=lambda t: (t[0], t[2]))
    return out


def _row_vals(grid, r, c0, c1):
    return [v for v in grid[r][c0:c1] if v is not None]


def _header_like(vals, width) -> bool:
    if not vals:
        return False
    strs = sum(isinstance(v, str) for v in vals)
    return len(vals) >= max(2, math.ceil(0.6 * width)) and strs / len(vals) >= 0.5


def _glue_fragments(grid, regions):
    """Merge regions that are really the continuation of the table above."""
    acc = []
    for r0, r1, c0, c1 in regions:
        vals = None
        for r in range(r0, r1):
            vals = _row_vals(grid, r, c0, c1)
            if vals:
                break
        w = c1 - c0
        target = None
        if vals and len(vals) >= 2:
            best = -1
            for a in acc:
                ov = min(a[3], c1) - max(a[2], c0)
                if ov <= 0 or a[1] > r0 or ov / (a[3] - a[2]) < 0.5 or ov / w < 0.5:
                    continue
                if a[1] > best:
                    best, target = a[1], a
            if target is not None:
                first = next(r for r in range(target[0], target[1]) if _row_vals(grid, r, target[2], target[3]))
                th = [norm(v) for v in _row_vals(grid, first, target[2], target[3])]
                rep = [norm(v) for v in vals] == th
                if not (rep or not _header_like(vals, w)):
                    target = None
        if target is not None:
            target[1] = r1
            target[2] = min(target[2], c0)
            target[3] = max(target[3], c1)
        else:
            acc.append([r0, r1, c0, c1])
    return [tuple(a) for a in acc]


def parse_region(sheet: SheetData, reg, opts: Options):
    """-> (Table | None, note | None)"""
    grid = sheet.grid
    r0, r1, c0, c1 = reg
    w = c1 - c0
    rows = [grid[r][c0:c1] for r in range(r0, r1)]
    fills = [sum(v is not None for v in row) for row in rows]
    maxfill = max(fills)
    if maxfill < 2:
        txt = next((str(v) for row in rows for v in row if v is not None), "")
        return None, f"note ignored at {get_column_letter(c0 + 1)}{r0 + 1}: {txt[:60]!r}"
    need = max(2, math.ceil(0.6 * maxfill))

    def is_title(row):
        vals = [v for v in row if v is not None]
        return len(vals) >= 2 and w >= 3 and len({str(v) for v in vals}) == 1

    h = None
    for i, row in enumerate(rows):
        vals = [v for v in row if v is not None]
        if fills[i] >= need and not is_title(row) and sum(isinstance(v, str) for v in vals) / len(vals) >= 0.5:
            h = i
            break
    if h is None:
        h = next((i for i, f in enumerate(fills) if f >= need), 0)

    title = None
    for row in reversed(rows[:h]):
        vals = [v for v in row if v is not None]
        if vals and len({str(v) for v in vals}) == 1:
            title = str(vals[0])
            break

    def row_has_merge(i):
        return any((r0 + i, c0 + j) in sheet.merged for j in range(w))

    h_end = h + 1
    if row_has_merge(h):
        while (h_end - h < 3 and h_end < len(rows) and fills[h_end] >= need
               and not any(_is_num_or_date(v) for v in rows[h_end] if v is not None)):
            h_end += 1

    header = []
    for j in range(w):
        parts = []
        for k in range(h, h_end):
            v = rows[k][j]
            if v is not None:
                t = _hdr_text(v)
                if not parts or parts[-1] != t:
                    parts.append(t)
        header.append(" | ".join(parts) if parts else None)
    hnorm = [norm(x) if x else "" for x in header]
    n_hdr = sum(1 for x in hnorm if x)

    data, nums = [], []
    for i in range(h_end, len(rows)):
        row = rows[i]
        vals = [v for v in row if v is not None]
        if not vals:
            continue
        if opts.skip_totals and isinstance(vals[0], str) and TOTAL_RE.match(vals[0]):
            continue
        if len(vals) >= 2 and sum(1 for j in range(w) if row[j] is not None and hnorm[j] and norm(row[j]) == hnorm[j]) >= max(2, 0.8 * n_hdr):
            continue
        data.append(list(row))
        nums.append(r0 + i + 1)
    while data and w >= 3 and sum(v is not None for v in data[-1]) == 1 and any(isinstance(v, str) for v in data[-1]):
        data.pop()
        nums.pop()

    keep = [j for j in range(w) if header[j] is not None or any(r[j] is not None for r in data)]
    if len(keep) < 2 or len(data) < opts.min_data_rows:
        return None, f"too small to be a table at {get_column_letter(c0 + 1)}{r0 + 1} ({len(data)} data rows)"

    names, norms, seen = [], [], {}
    for j in keep:
        name = header[j] or f"Column_{get_column_letter(c0 + j + 1)}"
        n = norm(name)
        k = seen.get(n, 0) + 1
        seen[n] = k
        if k > 1:
            name, n = f"{name} ({k})", f"{n}#{k}"
        names.append(name)
        norms.append(n)
    return Table(names, norms, [[r[j] for j in keep] for r in data], nums, r0, r1, c0, c1, title, r0 + h + 1), None


def analyze_sheet(sheet: SheetData, opts: Options):
    regions = _glue_fragments(sheet.grid, find_regions(sheet.grid, opts.row_gap, opts.col_gap))
    tables, notes = [], []
    for reg in regions:
        t, note = parse_region(sheet, reg, opts)
        if t:
            tables.append(t)
        elif note:
            notes.append(note)
    return tables, notes


# --------------------------------------------------------------------------
# matching tables across files
# --------------------------------------------------------------------------
def _jacc(a: set, b: set) -> float:
    return len(a & b) / len(a | b) if (a or b) else 0.0


@dataclass
class Chunk:
    file: str
    idx_map: list
    rows: list
    rownums: list


class Template:
    def __init__(self, table: Table):
        self.cols, self.index = [], {}
        self.canon = set(table.norms)
        self.union = set()
        self.chunks, self.titles, self.files = [], [], set()

    def score(self, table: Table) -> float:
        s = set(table.norms)
        return max(_jacc(s, self.canon), _jacc(s, self.union))

    def add(self, table: Table, fname: str):
        idx = []
        for n, h in zip(table.norms, table.header):
            if n not in self.index:
                self.index[n] = len(self.cols)
                self.cols.append(h)
            idx.append(self.index[n])
        self.union.update(table.norms)
        self.chunks.append(Chunk(fname, idx, table.rows, table.rownums))
        self.titles.append(table.title)
        self.files.add(fname)

    @property
    def nrows(self):
        return sum(len(c.rows) for c in self.chunks)

    @property
    def title(self):
        if self.titles and all(self.titles) and len({norm(t) for t in self.titles}) == 1:
            return self.titles[0]
        return None


def assign_tables(tables, templates, thr):
    pairs = [(t_.score(tb), ti, ki) for ti, tb in enumerate(tables) for ki, t_ in enumerate(templates)]
    pairs.sort(key=lambda p: (-p[0], p[1], p[2]))
    eff = 0.3 if len(tables) == 1 and len(templates) == 1 else thr
    res, used_t, used_k = {}, set(), set()
    for s, ti, ki in pairs:
        if s < eff:
            break
        if ti in used_t or ki in used_k:
            continue
        res[ti] = ki
        used_t.add(ti)
        used_k.add(ki)
    return res


@dataclass
class Group:
    display: str
    templates: list = field(default_factory=list)
    files: set = field(default_factory=set)


# --------------------------------------------------------------------------
# discovery
# --------------------------------------------------------------------------
def _natkey(s):
    return [int(p) if p.isdigit() else p.lower() for p in re.split(r"(\d+)", str(s))]


def _files_in(folder: Path, opts: Options, recursive=True):
    it = folder.rglob("*") if recursive else folder.glob("*")
    return sorted((p for p in it if p.is_file() and p.suffix.lower() in opts.extensions and not p.name.startswith(("~$", "."))),
                  key=lambda p: _natkey(p.relative_to(folder).as_posix()))


def discover(root, opts: Optional[Options] = None):
    opts = opts or Options()
    root = Path(root)
    months = []
    for mdir in sorted((d for d in root.iterdir() if d.is_dir() and not d.name.startswith((".", "_"))), key=lambda d: _natkey(d.name)):
        types = []
        loose = _files_in(mdir, opts, recursive=False)
        if loose:
            types.append({"name": "(files)", "path": str(mdir), "files": [str(f) for f in loose]})
        for tdir in sorted((d for d in mdir.iterdir() if d.is_dir() and not d.name.startswith((".", "_"))), key=lambda d: _natkey(d.name)):
            fl = _files_in(tdir, opts)
            if fl:
                types.append({"name": tdir.name, "path": str(tdir), "files": [str(f) for f in fl]})
        if types:
            months.append({"name": mdir.name, "path": str(mdir), "types": types})
    return months


# --------------------------------------------------------------------------
# preview
# --------------------------------------------------------------------------
def _js(v):
    if v is None or isinstance(v, (bool, int, float)):
        return v
    if isinstance(v, (dt.datetime, dt.date, dt.time)):
        return v.isoformat()
    return str(v)


def analyze_file(path, opts: Optional[Options] = None):
    opts = opts or Options()
    out = []
    for sh in read_sheets(Path(path), opts):
        tables, notes = analyze_sheet(sh, opts)
        out.append({
            "sheet": sh.name,
            "hidden": sh.hidden,
            "dims": f"{len(sh.grid)} rows x {len(sh.grid[0]) if sh.grid else 0} cols",
            "tables": [{
                "ref": t.ref, "title": t.title, "header_row": t.header_row,
                "columns": t.header, "n_rows": len(t.rows),
                "sample": [[_js(v) for v in r] for r in t.rows[:5]],
            } for t in tables],
            "notes": notes,
        })
    return out


# --------------------------------------------------------------------------
# consolidation
# --------------------------------------------------------------------------
def _safe_part(s):
    return re.sub(r"[\[\]:*?/\\]", "_", str(s)).strip().strip("'")


def unique_sheet_name(used: set, type_, sheet, suffix="", with_type=True):
    type_, sheet = _safe_part(type_), _safe_part(sheet)
    allowed = 31 - len(suffix)
    stem = f"{type_}_{sheet}" if with_type else sheet
    if len(stem) > allowed:
        if with_type:
            tl = min(len(type_), max(4, allowed // 3))
            stem = type_[:tl] + "_" + sheet[: allowed - tl - 1]
        else:
            stem = stem[:allowed]
    name = (stem + suffix) or "Sheet"
    n = 1
    while name.lower() in used:
        n += 1
        tail = f"~{n}"
        name = (stem + suffix)[: 31 - len(tail)] + tail
    used.add(name.lower())
    return name


_HDR_FONT = Font(bold=True, color="FFFFFF")
_HDR_FILL = PatternFill("solid", fgColor="1F4E78")
_TTL_FONT = Font(bold=True, size=12)
_OK_TYPES = (str, int, float, bool, dt.date, dt.datetime, dt.time, dt.timedelta, Decimal)


def _wcell(ws, v, font=None, fill=None):
    c = WriteOnlyCell(ws, value=v)
    if v.__class__ is str and v.startswith("="):
        c.data_type = "s"
    if font:
        c.font = font
    if fill:
        c.fill = fill
    return c


def _val(v):
    if v is None:
        return None
    if not isinstance(v, _OK_TYPES):
        return str(v)
    if v.__class__ is str and v.startswith("="):
        return v
    return v


def _plain(ws, v):
    """Convert a data value to something append() accepts (wrapping formulas-as-text)."""
    v = _val(v)
    if v.__class__ is str and v.startswith("="):
        return _wcell(ws, v)
    return v


class BookWriter:
    def __init__(self):
        self.wb = Workbook(write_only=True)
        self.used = set()
        self.sheets = {}

    def add_sheet(self, name):
        ws = self.wb.create_sheet(title=name)
        self.used.add(name.lower())
        return ws

    def save(self, path):
        self.wb.save(path)
        self.wb.close()


def _table_rows(tpl: Template, opts: Options):
    """Yield output rows (lists) for a template, with optional source columns."""
    n = len(tpl.cols)
    for ch in tpl.chunks:
        im = ch.idx_map
        for row, rn in zip(ch.rows, ch.rownums):
            out = [None] * n
            for j, i in enumerate(im):
                out[i] = row[j]
            pre = []
            if opts.add_source_file:
                pre.append(ch.file)
            if opts.add_source_row:
                pre.append(rn)
            yield pre + out if pre else out


def _out_headers(tpl: Template, opts: Options):
    pre = (["Source File"] if opts.add_source_file else []) + (["Source Row"] if opts.add_source_row else [])
    return pre + list(tpl.cols)


def _write_plan(book: BookWriter, plan, opts: Options, with_type: bool):
    """plan: list of sheet entries {name, items:[(label, Template)]}; writes data sheets."""
    for entry in plan:
        stacked = len(entry["items"]) > 1 and opts.multi_table_layout == "stacked"
        if stacked or len(entry["items"]) == 1:
            groups = [entry["items"]] if stacked else [[it] for it in entry["items"]]
        else:
            groups = [[it] for it in entry["items"]]
        for gi, items in enumerate(groups):
            name = entry["names"][gi]
            ws = book.add_sheet(name)
            base_stem = name
            sheet_no = 1
            # column widths = max over tables on this sheet
            maxw = {}
            for _, tpl in items:
                hdrs = _out_headers(tpl, opts)
                sample = []
                for k, r in enumerate(_table_rows(tpl, opts)):
                    sample.append(r)
                    if k >= 100:
                        break
                for j, h in enumerate(hdrs):
                    m = len(str(h))
                    for r in sample:
                        if r[j] is not None:
                            m = max(m, min(len(str(r[j])), 60))
                    maxw[j] = max(maxw.get(j, 0), min(60, max(10, m + 2)))
            for j, wd in maxw.items():
                ws.column_dimensions[get_column_letter(j + 1)].width = wd
            if not stacked:
                ws.freeze_panes = "A2"
            cur = 0
            for label, tpl in items:
                hdrs = _out_headers(tpl, opts)
                if stacked:
                    ws.append([_wcell(ws, label, font=_TTL_FONT)])
                    cur += 1
                hdr_cells = [_wcell(ws, h, font=_HDR_FONT, fill=_HDR_FILL) for h in hdrs]
                ws.append(hdr_cells)
                cur += 1
                for row in _table_rows(tpl, opts):
                    if cur >= EXCEL_ROW_LIMIT:
                        sheet_no += 1
                        ws = book.add_sheet(unique_sheet_name(book.used, "", base_stem, f" #{sheet_no}", with_type=False))
                        if not stacked:
                            ws.freeze_panes = "A2"
                        ws.append([_wcell(ws, h, font=_HDR_FONT, fill=_HDR_FILL) for h in hdrs])
                        cur = 1
                    ws.append([_plain(ws, v) for v in row])
                    cur += 1
                if stacked:
                    ws.append([])
                    cur += 1


def _plain_sheet(book, name, header, rows, widths=None):
    ws = book.add_sheet(name)
    for j, w in enumerate(widths or []):
        ws.column_dimensions[get_column_letter(j + 1)].width = w
    ws.freeze_panes = "A2"
    ws.append([_wcell(ws, h, font=_HDR_FONT, fill=_HDR_FILL) for h in header])
    for r in rows:
        ws.append([_plain(ws, v) for v in r])


def _process_type(files, opts: Options, issues, log, stop, tick, type_name):
    groups: "OrderedDict[str, Group]" = OrderedDict()
    ok = 0
    for fp in files:
        if stop and stop():
            raise Cancelled()
        fp = Path(fp)
        try:
            sheets = read_sheets(fp, opts)
        except Exception as e:  # noqa: BLE001 - one bad file must not kill the batch
            issues.append(("ERROR", type_name, fp.name, f"could not read file: {e}"))
            log("ERROR", f"{type_name}/{fp.name}: could not read ({e})")
            tick()
            continue
        ok += 1
        n_tables = 0
        for sh in sheets:
            if sh.hidden and opts.skip_hidden_sheets:
                continue
            tables, notes = analyze_sheet(sh, opts)
            for n in notes:
                issues.append(("INFO", type_name, f"{fp.name} / {sh.name}", n))
            if not tables:
                issues.append(("WARN", type_name, f"{fp.name} / {sh.name}", "no table detected in sheet"))
                continue
            key = re.sub(r"\s+", " ", sh.name.strip().lower())
            g = groups.setdefault(key, Group(sh.name.strip()))
            assignment = assign_tables(tables, g.templates, opts.match_threshold)
            for ti, tb in enumerate(tables):
                if ti not in assignment:
                    g.templates.append(Template(tb))
                    assignment[ti] = len(g.templates) - 1
                g.templates[assignment[ti]].add(tb, fp.name)
            g.files.add(fp.name)
            n_tables += len(tables)
        log("INFO", f"{type_name}/{fp.name}: {len(sheets)} sheet(s), {n_tables} table(s)")
        tick()
    names = [Path(f).name for f in files]
    for g in groups.values():
        absent = [n for n in names if n not in g.files]
        if absent:
            issues.append(("WARN", type_name, g.display, f"sheet missing in {len(absent)} file(s): {', '.join(absent[:8])}{' ...' if len(absent) > 8 else ''}"))
        for i, t in enumerate(g.templates, 1):
            if len(t.files) < len(g.files):
                issues.append(("WARN", type_name, f"{g.display} / T{i}", f"table found in only {len(t.files)} of {len(g.files)} file(s) having this sheet"))
    return groups, ok


def _safe_file(s):
    return re.sub(r'[<>:"/\\|?*]+', "_", str(s)).strip() or "unnamed"


def consolidate(root, out_dir, months=None, opts=None, log: Callable = None, progress: Callable = None, stop: Callable = None):
    """Consolidate selected months (default: all). Returns a result dict."""
    opts = opts if isinstance(opts, Options) else Options.from_dict(opts)
    log = log or (lambda lvl, msg: print(f"[{lvl}] {msg}"))
    root, out_dir = Path(root), Path(out_dir)
    out_dir.mkdir(parents=True, exist_ok=True)
    found = [m for m in discover(root, opts) if not months or m["name"] in months]
    total = sum(len(t["files"]) for m in found for t in m["types"])
    done = [0]

    def tick():
        done[0] += 1
        if progress:
            progress(done[0], total)

    results = []
    for m in found:
        if stop and stop():
            raise Cancelled()
        log("INFO", f"=== Month {m['name']} ===")
        issues = []
        type_groups = OrderedDict()
        for t in m["types"]:
            log("INFO", f"-- {t['name']}: {len(t['files'])} file(s)")
            groups, ok = _process_type(t["files"], opts, issues, log, stop, tick, t["name"])
            type_groups[t["name"]] = (groups, ok, len(t["files"]))

        def build(book: BookWriter, tnames, with_type):
            index_rows, plan = [], []
            for tn in tnames:
                groups, ok, nfiles = type_groups[tn]
                for g in groups.values():
                    items = []
                    for i, tpl in enumerate(g.templates, 1):
                        label = f"T{i}" + (f" - {tpl.title}" if tpl.title else "")
                        items.append((f"{g.display} / {label}" if len(g.templates) > 1 else g.display, tpl, label))
                    multi = len(g.templates) > 1
                    stacked = multi and opts.multi_table_layout == "stacked"
                    names = []
                    if not multi or stacked:
                        names.append(unique_sheet_name(book.used, tn, g.display, "", with_type))
                    else:
                        for (_, tpl, label) in items:
                            sfx = f" - {label.split(' - ')[0]}"
                            names.append(unique_sheet_name(book.used, tn, g.display, sfx, with_type))
                    plan.append({"names": names, "items": [(lbl, tpl) for lbl, tpl, _ in items]})
                    if not multi or stacked:
                        for lbl, tpl, label in items:
                            index_rows.append([names[0], tn, g.display, label if multi else "", len(g.files), tpl.nrows, len(tpl.cols)])
                    else:
                        for nm, (lbl, tpl, label) in zip(names, items):
                            index_rows.append([nm, tn, g.display, label, len(g.files), tpl.nrows, len(tpl.cols)])
            book.used.add("_index")
            book.used.add("_log")
            _plain_sheet(book, "_Index", ["Output sheet", "Data type", "Source sheet", "Table", "Files contributing", "Rows", "Columns"],
                         index_rows, [28, 22, 28, 28, 18, 12, 10])
            _plain_sheet(book, "_Log", ["Level", "Data type", "Where", "Message"],
                         [list(i) for i in issues if i[1] in tnames or not i[1]], [10, 22, 40, 100])
            _write_plan(book, plan, opts, with_type)
            return sum(1 for r in index_rows), sum(r[5] for r in index_rows)

        outputs = []
        mname = _safe_file(m["name"])
        if opts.per_type_files:
            for tn in type_groups:
                book = BookWriter()
                ns, nr = build(book, [tn], False)
                path = out_dir / f"{mname}_{_safe_file(tn)}_consolidated.xlsx"
                book.save(path)
                outputs.append({"path": str(path), "sheets": ns, "rows": nr})
                log("OK", f"wrote {path.name}: {ns} consolidated table(s), {nr} rows")
        else:
            book = BookWriter()
            ns, nr = build(book, list(type_groups), True)
            path = out_dir / f"{mname}_consolidated.xlsx"
            book.save(path)
            outputs.append({"path": str(path), "sheets": ns, "rows": nr})
            log("OK", f"wrote {path.name}: {ns} consolidated table(s), {nr} rows")
        results.append({"month": m["name"], "outputs": outputs,
                        "issues": [{"level": a, "type": b, "where": c, "msg": d} for a, b, c, d in issues if a != "INFO"]})
    return {"months": results, "files": total}

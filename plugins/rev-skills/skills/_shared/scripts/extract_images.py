"""Pictures pasted on design-doc sheets: extract them as PNG, and list the live cells they cover.

Reads the .xlsx as a zip only (no openpyxl load — that takes 10+ minutes on a merge-heavy workbook such
as PSJAO401; never saves anything). For every picture anchored on a visible sheet (skips 詳細設計* /
*画面ｲﾒｰｼﾞ* sheets, and hidden sheets unless --hidden) writes into OUT_DIR:
  <sheet>_r<row>c<col>_<media>.png   the picture (EMF/WMF rendered with Pillow; needs Windows GDI)
  index.tsv     sheet, kind, from [r,c], to [r,c], media, format, px WxH, png, est_tokens, note, covered
  covered.txt   "<sheet> [r,c] <live text>  (under <png>)" for every live cell lying under a picture
                (more than half of an edge row/column covered) — text nobody can see on the sheet;
                the revision margin (col >= --margin-col, default 105 = DA列) is left out.
                Needs --dump: the folder of the live text dump (live_dump.py or the COM dump,
                <prefix>_<sheet>.txt files of [r,c]=text tokens); without it covered is "-".
kind = 画面ﾚｲｱｳﾄ for a 画面設計書 picture anchored above the sheet's 「Ⅱ．」 heading (found in the dump, or
read from the zip with xlsx_zip.py when there is none), else その他. note = icon (long side < 100 px, a button glyph — don't read), same as <png> (identical
bytes, already listed), unreadable (…), or empty. est_tokens ~ px area / 750 after scaling the long
side to <= 1568 px.

usage: python extract_images.py <workbook.xlsx> <out_dir> [--dump DIR] [--sheets REGEX] [--hidden] [--margin-col N]
"""
import argparse, glob, hashlib, html, io, os, posixpath, re, sys, zipfile
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from xlsx_zip import iter_sheets   # live text from the zip, used for the 「Ⅱ．」 heading when there is no dump

EMU_PX = 9525

def rels(z, part):
    p = posixpath.join(posixpath.dirname(part), "_rels", posixpath.basename(part) + ".rels")
    if p not in z.namelist(): return {}
    out = {}
    for t in re.findall(r"<Relationship\b[^>]*>", z.read(p).decode("utf-8")):
        i = re.search(r'Id="([^"]+)"', t); g = re.search(r'Target="([^"]+)"', t)
        if i and g:
            tgt = g.group(1)
            out[i.group(1)] = tgt.lstrip("/") if tgt.startswith("/") else posixpath.normpath(posixpath.join(posixpath.dirname(part), tgt))
    return out

class Grid:
    """row heights / column widths in px, from the sheet XML (cheap regex, no cell parsing)"""
    def __init__(self, xml):
        head = xml.split("<sheetData", 1)[0]
        f = re.search(r"<sheetFormatPr\b[^>]*>", head)
        dh = re.search(r'defaultRowHeight="([\d.]+)"', f.group(0)) if f else None
        dw = re.search(r'defaultColWidth="([\d.]+)"', f.group(0)) if f else None
        self.dh = float(dh.group(1)) if dh else 15.0
        self.dw = float(dw.group(1)) if dw else 8.43
        self.cols = []
        for t in re.findall(r"<col\b[^>]*>", head):
            lo, hi, w = (re.search(rf'{k}="([\d.]+)"', t) for k in ("min", "max", "width"))
            if lo and hi and w: self.cols.append((int(lo.group(1)), int(hi.group(1)), float(w.group(1))))
        self.rows = {int(r): float(h) for r, h in re.findall(r'<row\b[^>]*?\br="(\d+)"[^>]*?\bht="([\d.]+)"', xml)}
    def row_px(self, r): return self.rows.get(r, self.dh) * 96 / 72
    def col_px(self, c):
        for lo, hi, w in self.cols:
            if lo <= c <= hi: return w * 7 + 5
        return self.dw * 7 + 5

def walk(start, offset_px, size_px, step):
    i, left = start, offset_px + size_px
    while i < start + 5000:
        s = step(i)
        if left <= s: return i, left
        left -= s; i += 1
    return i, 0

def anchors(dx):
    for m in re.finditer(r"<xdr:(twoCellAnchor|oneCellAnchor|absoluteAnchor)\b.*?</xdr:\1>", dx, re.S):
        a = m.group(0)
        emb = re.search(r'r:embed="([^"]+)"', a)
        if not emb: continue                       # shapes, charts, text boxes: not pictures
        def cell(tag):
            t = re.search(rf"<xdr:{tag}>(.*?)</xdr:{tag}>", a, re.S)
            if not t: return None
            g = lambda k: int(re.search(rf"<xdr:{k}>(-?\d+)</xdr:{k}>", t.group(1)).group(1))
            return g("row") + 1, g("col") + 1, g("rowOff") / EMU_PX, g("colOff") / EMU_PX
        ext = re.search(r'<xdr:ext cx="(\d+)" cy="(\d+)"', a)
        yield m.group(1), emb.group(1), cell("from"), cell("to"), \
            (int(ext.group(1)) / EMU_PX, int(ext.group(2)) / EMU_PX) if ext else None

def load_dump(dump_dir, sheet):
    """{(r,c): text} from <prefix>_<sheet>.txt in the dump folder, or None if not found"""
    safe = "".join("_" if ch in '\\/:*?"<>|' else ch for ch in sheet)
    hits = [p for p in glob.glob(os.path.join(dump_dir, "*.txt")) if os.path.basename(p).endswith("_" + safe + ".txt")]
    if not hits: return None
    cells = {}
    for line in open(hits[0], encoding="utf-8"):
        for m in re.finditer(r"\[(\d+),(\d+)\]=(.*?)(?= \| \[\d+,\d+\]=|$)", line.rstrip("\n")):
            cells[(int(m.group(1)), int(m.group(2)))] = m.group(3)
    return cells

def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("workbook"); ap.add_argument("out_dir")
    ap.add_argument("--dump"); ap.add_argument("--sheets"); ap.add_argument("--hidden", action="store_true")
    ap.add_argument("--margin-col", type=int, default=105,
                    help="cells from this column on are the revision margin and never reported as covered")
    a = ap.parse_args()
    os.makedirs(a.out_dir, exist_ok=True)
    from PIL import Image
    z = zipfile.ZipFile(a.workbook)
    wbx = z.read("xl/workbook.xml").decode("utf-8")
    wrels = rels(z, "xl/workbook.xml")
    index = ["sheet\tkind\tfrom\tto\tmedia\tformat\tpx\tpng\test_tokens\tnote\tcovered"]
    covered_lines = []; seen = {}; total_tokens = 0; missing_dump = []
    for t in re.findall(r"<sheet\b[^>]*>", wbx):
        name = html.unescape(re.search(r'name="([^"]+)"', t).group(1))
        if re.search(r'state="(hidden|veryHidden)"', t) and not a.hidden: continue
        if name.startswith("詳細設計") or "画面ｲﾒｰｼﾞ" in name: continue
        if a.sheets and not re.search(a.sheets, name): continue
        sheet_part = wrels.get(re.search(r'r:id="([^"]+)"', t).group(1))
        if not sheet_part or sheet_part not in z.namelist(): continue
        drawings = [p for p in rels(z, sheet_part).values() if re.search(r"drawings/drawing\d+\.xml$", p)]
        if not drawings: continue
        grid = Grid(z.read(sheet_part).decode("utf-8"))
        cells = load_dump(a.dump, name) if a.dump else None
        if a.dump and cells is None: missing_dump.append(name)
        sec2 = None
        if "画面設計書" in name:
            if cells is not None:
                hs = [r for (r, c), v in cells.items() if c <= 6 and re.match(r"\s*Ⅱ[．.]", v)]
            else:                                   # no dump: read the heading columns from the zip
                hs = [r for _, _, rows in iter_sheets(a.workbook, name_filter=re.compile("^" + re.escape(name) + "$"), max_col=6)
                      for r, cs in rows.items() for v in cs.values() if re.match(r"\s*Ⅱ[．.]", v)]
            sec2 = min(hs) if hs else None
        for dp in drawings:
            drels = rels(z, dp)
            for typ, emb, fr, to, ext in anchors(z.read(dp).decode("utf-8")):
                media = drels.get(emb)
                if not media or media not in z.namelist(): continue
                data = z.read(media)
                if fr is None: fr = (1, 1, 0, 0)
                if to is None and ext:                  # oneCellAnchor: walk the grid by the extent
                    r1, ro = walk(fr[0], fr[2], ext[1], grid.row_px); c1, co = walk(fr[1], fr[3], ext[0], grid.col_px)
                    to = (r1, c1, ro, co)
                if to is None: to = fr
                safe = "".join("_" if ch in '\\/:*?"<>|' else ch for ch in name)
                png = f"{safe}_r{fr[0]}c{fr[1]}_{posixpath.splitext(posixpath.basename(media))[0]}.png"
                h = hashlib.sha1(data).hexdigest(); note = ""; size = ""; est = 0
                try:
                    im = Image.open(io.BytesIO(data)); im.load()
                    w, hgt = im.size; size = f"{w}x{hgt}"
                    if max(w, hgt) < 100: note = "icon"
                    elif h in seen: note = f"same as {seen[h]}"
                    else:
                        seen[h] = png
                        if im.mode not in ("RGB", "RGBA", "L"): im = im.convert("RGBA")
                        im.save(os.path.join(a.out_dir, png))
                        k = min(1.0, 1568 / max(w, hgt)); est = int(w * k * hgt * k / 750); total_tokens += est
                except Exception as e:
                    note = f"unreadable ({type(e).__name__})"
                kind = "画面ﾚｲｱｳﾄ" if ("画面設計書" in name and (sec2 is None or fr[0] < sec2)) else "その他"
                ncov = "-"
                if cells is not None:
                    # an edge row/column counts only when the picture spans more than half of it
                    r0 = fr[0] + (fr[2] > grid.row_px(fr[0]) / 2); c0 = fr[1] + (fr[3] > grid.col_px(fr[1]) / 2)
                    r1 = to[0] - (to[2] < grid.row_px(to[0]) / 2); c1 = to[1] - (to[3] < grid.col_px(to[1]) / 2)
                    ncov = 0
                    for (r, c), v in sorted(cells.items()):
                        if c >= a.margin_col: continue      # revision margin (DA列~): change notes, not text
                        if r0 <= r <= r1 and c0 <= c <= c1 and v.strip():
                            ncov += 1
                            covered_lines.append(f"{name} [{r},{c}] {v.strip()[:200]}  (under {png})")
                index.append("\t".join(map(str, [name, kind, f"[{fr[0]},{fr[1]}]", f"[{to[0]},{to[1]}]",
                                                 posixpath.basename(media), posixpath.splitext(media)[1].lstrip(".").lower(),
                                                 size, "-" if note else png, est, note, ncov])))
    def key(line):                                  # sheet order as in the workbook, then anchor row/col
        f = line.split("\t"); r, c = map(int, f[2].strip("[]").split(","))
        return order.index(f[0]), r, c
    order = list(dict.fromkeys(l.split("\t")[0] for l in index[1:]))
    index = index[:1] + sorted(index[1:], key=key)
    open(os.path.join(a.out_dir, "index.tsv"), "w", encoding="utf-8").write("\n".join(index) + "\n")
    open(os.path.join(a.out_dir, "covered.txt"), "w", encoding="utf-8").write("\n".join(covered_lines) + ("\n" if covered_lines else ""))
    print(f"{len(index) - 1} pictures, {len(seen)} written, est_tokens={total_tokens}, "
          f"covered_cells={len(covered_lines) if a.dump else '-'}")
    if missing_dump:
        print("WARNING: no dump file for sheet(s) with pictures: " + ", ".join(missing_dump), file=sys.stderr)

if __name__ == "__main__":
    main()

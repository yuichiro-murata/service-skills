"""Read live cell text straight from an .xlsx zip — no openpyxl load, no Excel, never writes.

Same live rule as live_dump.py at cell level: a cell whose font is struck or gray (r = g = b, 80 < r < 220)
is dead; inside rich text a run with its own rPr overrides the cell font (struck run dead), a bare run
inherits it. Merged non-anchor cells are skipped, as in live_dump.py: Excel can keep a stale value there
(SXJCB147 G704:P704 all hold YOTO; 8,211 such values across the 122 工程管理 books). Time-of-day cells come back
as the stored day fraction (0.5), where live_dump prints 12:00:00 — no number formats are read.
Used by build_param_index.py and build_zoom_list.py; sheet_fills.py carries its own copy (it also needs
fills). Fast on sheets with a runaway used range (03.ｽﾞｰﾑ管理台帳_基準情報's 1,048,566 rows): only
stored cells are visited.

    from xlsx_zip import iter_sheets
    for name, state, rows in iter_sheets(path, name_filter=re.compile("^ｽﾞｰﾑ一覧"), max_col=100):
        rows -> {row: {col: text}}
"""
import html, re, zipfile

T = re.compile(r"<t[^>]*>(.*?)</t>", re.S)
STRIKE = re.compile(r'<strike(?! val="0")')


def _txt(s):
    return html.unescape("".join(T.findall(re.sub(r"<rPh.*?</rPh>", "", s, flags=re.S))))


def _col(letters):
    n = 0
    for ch in letters:
        n = n * 26 + ord(ch) - 64
    return n


def _dead_font(f):
    if STRIKE.search(f):
        return True
    m = re.search(r'<color [^>]*rgb="[0-9A-Fa-f]{2}([0-9A-Fa-f]{6})"', f)
    if m:
        h = m.group(1)
        r, g, b = int(h[0:2], 16), int(h[2:4], 16), int(h[4:6], 16)
        return r == g == b and 80 < r < 220
    return False


def _runs(si):
    out = []
    for r in re.findall(r"<r>(.*?)</r>", si, re.S):
        rpr = re.search(r"<rPr>(.*?)</rPr>", r, re.S)
        out.append((_txt(re.sub(r"<rPr>.*?</rPr>", "", r, flags=re.S)), bool(rpr), bool(rpr and STRIKE.search(rpr.group(1)))))
    return out


def _live(entry, cell_dead):
    plain, runs = entry
    if not runs:
        return "" if cell_dead else plain
    return "".join(t for t, own, st in runs if (not st if own else not cell_dead))


def iter_sheets(path, name_filter=None, include_hidden=False, max_col=None):
    """yield (name, state, rows{r: {c: text}}) — live text only."""
    z = zipfile.ZipFile(path)
    try:
        sx = z.read("xl/sharedStrings.xml").decode("utf-8")
    except KeyError:
        sx = ""
    ss = [(_txt(si), _runs(si)) for si in re.findall(r"<si>(.*?)</si>", sx, re.S)]
    st = z.read("xl/styles.xml").decode("utf-8")
    fonts = re.search(r"<fonts[^>]*>(.*?)</fonts>", st, re.S).group(1)
    fdead = [_dead_font(f) for f in re.findall(r"<font(?:/>|>.*?</font>)", fonts, re.S)]
    xs = []
    for m in re.finditer(r"<xf ([^>]*?)/?>", re.search(r"<cellXfs[^>]*>(.*?)</cellXfs>", st, re.S).group(1)):
        f = re.search(r'fontId="(\d+)"', m.group(1))
        f = int(f.group(1)) if f else 0
        xs.append(fdead[f] if f < len(fdead) else False)
    wb = z.read("xl/workbook.xml").decode("utf-8")
    rels = z.read("xl/_rels/workbook.xml.rels").decode("utf-8")
    rid = {re.search(r'Id="([^"]+)"', m.group(0)).group(1): re.search(r'Target="([^"]+)"', m.group(0)).group(1)
           for m in re.finditer(r"<Relationship [^>]*>", rels)}
    for m in re.finditer(r"<sheet ([^>]*)/>", wb):
        a = m.group(1)
        state = (re.search(r'state="([^"]*)"', a) or [None, "visible"])[1]
        if state != "visible" and not include_hidden:
            continue
        name = html.unescape(re.search(r'name="([^"]*)"', a).group(1))
        if name_filter and not name_filter.search(name):
            continue
        tg = rid[re.search(r'r:id="([^"]+)"', a).group(1)].lstrip("/")
        tg = tg if tg.startswith("xl/") else "xl/" + tg
        x = z.read(tg).decode("utf-8")
        skip = set()
        for mm in re.finditer(r'<mergeCell ref="([A-Z]+)(\d+):([A-Z]+)(\d+)"', x):
            c1, r1, c2, r2 = _col(mm.group(1)), int(mm.group(2)), _col(mm.group(3)), int(mm.group(4))
            if (r2 - r1 + 1) * (c2 - c1 + 1) <= 200000:
                skip.update((r, c) for r in range(r1, r2 + 1) for c in range(c1, c2 + 1) if (r, c) != (r1, c1))
        rows = {}
        for c in re.finditer(r'<c r="([A-Z]+)(\d+)"([^>]*?)(?:/>|>(.*?)</c>)', x, re.S):
            body = c.group(4)
            if not body:
                continue
            col = _col(c.group(1))
            if max_col and col > max_col:
                continue
            if (int(c.group(2)), col) in skip:
                continue
            sm = re.search(r's="(\d+)"', c.group(3))
            dead = bool(sm and int(sm.group(1)) < len(xs) and xs[int(sm.group(1))])
            t = (re.search(r't="(\w+)"', c.group(3)) or [None, "n"])[1]
            v = re.search(r"<v(?: [^>]*)?>(.*?)</v>", body, re.S)   # <v xml:space="preserve"> occurs
            if t == "s" and v:
                txt = _live(ss[int(v.group(1))], dead)
            elif t == "inlineStr":
                txt = _live((_txt(body), _runs(body)), dead)
            elif v and not dead:
                txt = html.unescape(v.group(1))
            else:
                continue
            if txt.strip():
                rows.setdefault(int(c.group(2)), {})[col] = txt
        yield name, state, rows

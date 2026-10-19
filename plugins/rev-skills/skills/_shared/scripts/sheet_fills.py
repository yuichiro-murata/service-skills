"""Cell fills (and live text) of selected sheets, read from the xlsx zip — no openpyxl load, no Excel.

For design-doc-internal-consistency check 14: the field colours of a 画面ｲﾒｰｼﾞ sheet (white = input,
filled = display only, pink = required). Writes one line per cell that has a solid fill or text:

    [row,col]	fill	merge	text

fill is the fgColor as `RRGGBB`, `theme:N` / `theme:N tint:T` (theme colours are not resolved),
`indexed:N`, or `-` for no fill. merge is the merged range the cell anchors (`A1:D1`) or empty.
text is the cell's value with struck runs removed (the same live rule as live_dump.py: a whole-cell
strike or every run struck drops it). Read-only; never writes the workbook.

usage: python sheet_fills.py <workbook.xlsx> <sheet-regex> [out.tsv]
"""
import html, re, sys, zipfile

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
    """struck, or a gray rgb colour (r = g = b, 80 < r < 220) — live_dump.py's cell-level rule."""
    if STRIKE.search(f):
        return True
    m = re.search(r'<color [^>]*rgb="[0-9A-Fa-f]{2}([0-9A-Fa-f]{6})"', f)
    if m:
        h = m.group(1); r, g, b = int(h[0:2], 16), int(h[2:4], 16), int(h[4:6], 16)
        return r == g == b and 80 < r < 220
    return False


def _runs(si):
    """[(text, has_own_rPr, struck)] of one <si>; [] for a plain string. live_dump.py's rule: a run with its
    own rPr overrides the cell font, a bare run inherits it."""
    out = []
    for r in re.findall(r"<r>(.*?)</r>", si, re.S):
        rpr = re.search(r"<rPr>(.*?)</rPr>", r, re.S)
        out.append((_txt(re.sub(r"<rPr>.*?</rPr>", "", r, flags=re.S)), bool(rpr), bool(rpr and STRIKE.search(rpr.group(1)))))
    return out


def _live(si_entry, cell_struck):
    plain, runs = si_entry
    if not runs:
        return "" if cell_struck else plain
    return "".join(t for t, own, st in runs if (not st if own else not cell_struck))


def _shared(z):
    try:
        x = z.read("xl/sharedStrings.xml").decode("utf-8")
    except KeyError:
        return []
    return [(_txt(si), _runs(si)) for si in re.findall(r"<si>(.*?)</si>", x, re.S)]


def _styles(z):
    x = z.read("xl/styles.xml").decode("utf-8")
    fonts = re.search(r"<fonts[^>]*>(.*?)</fonts>", x, re.S).group(1)
    fstrike = [_dead_font(f) for f in re.findall(r"<font(?:/>|>.*?</font>)", fonts, re.S)]
    fills = []
    fx = re.search(r"<fills[^>]*>(.*?)</fills>", x, re.S)
    for f in re.findall(r"<fill(?:/>|>.*?</fill>)", fx.group(1) if fx else "", re.S):
        if not re.search(r'patternType="(solid|gray\d+|darkGray|mediumGray|lightGray)"', f):
            fills.append("-"); continue
        fg = re.search(r"<fgColor ([^>]*)/?>", f)
        a = fg.group(1) if fg else ""
        m = re.search(r'rgb="([0-9A-Fa-f]{6,8})"', a)
        if m:
            fills.append(m.group(1)[-6:].upper())
        elif "theme=" in a:
            th = re.search(r'theme="(\d+)"', a).group(1)
            tn = re.search(r'tint="([-0-9.Ee]+)"', a)
            fills.append(f"theme:{th}" + (f" tint:{float(tn.group(1)):.2f}" if tn else ""))
        elif "indexed=" in a:
            fills.append("indexed:" + re.search(r'indexed="(\d+)"', a).group(1))
        else:
            fills.append("solid")
    xfs = re.search(r"<cellXfs[^>]*>(.*?)</cellXfs>", x, re.S).group(1)
    res = []
    for m in re.finditer(r"<xf ([^>]*?)/?>", xfs):
        g = lambda k: int((re.search(k + r'="(\d+)"', m.group(1)) or [0, 0])[1])
        fid, fo = g("fillId"), g("fontId")
        res.append((fills[fid] if fid < len(fills) else "-", fstrike[fo] if fo < len(fstrike) else False))
    return res


def main():
    if len(sys.argv) < 3:
        sys.exit(__doc__)
    path, rx = sys.argv[1], re.compile(sys.argv[2])
    out = open(sys.argv[3], "w", encoding="utf-8") if len(sys.argv) > 3 else None
    if out is None:
        sys.stdout.reconfigure(encoding="utf-8")
    w = out.write if out else sys.stdout.write
    z = zipfile.ZipFile(path)
    wb = z.read("xl/workbook.xml").decode("utf-8")
    rels = z.read("xl/_rels/workbook.xml.rels").decode("utf-8")
    rid = {}
    for m in re.finditer(r"<Relationship [^>]*>", rels):
        i = re.search(r'Id="([^"]+)"', m.group(0)).group(1)
        rid[i] = re.search(r'Target="([^"]+)"', m.group(0)).group(1)
    ss, xs = _shared(z), _styles(z)
    hit = False
    for m in re.finditer(r"<sheet ([^>]*)/>", wb):
        name = html.unescape(re.search(r'name="([^"]*)"', m.group(1)).group(1))
        if not rx.search(name):
            continue
        hit = True
        tg = rid[re.search(r'r:id="([^"]+)"', m.group(1)).group(1)].lstrip("/")
        tg = tg if tg.startswith("xl/") else "xl/" + tg
        x = z.read(tg).decode("utf-8")
        merges = {}
        nonanchor = set()   # text there is a stale value Excel kept under the merge (live_dump drops it)
        for mm in re.finditer(r'<mergeCell ref="([A-Z]+)(\d+):([A-Z]+)(\d+)"', x):
            r1, c1, c2, r2 = int(mm.group(2)), _col(mm.group(1)), _col(mm.group(3)), int(mm.group(4))
            merges[(r1, c1)] = f"{mm.group(1)}{mm.group(2)}:{mm.group(3)}{mm.group(4)}"
            if (r2 - r1 + 1) * (c2 - c1 + 1) <= 200000:
                nonanchor.update((r, c) for r in range(r1, r2 + 1) for c in range(c1, c2 + 1) if (r, c) != (r1, c1))
        w(f"## {name}\n")
        for c in re.finditer(r'<c r="([A-Z]+)(\d+)"([^>]*?)(?:/>|>(.*?)</c>)', x, re.S):
            r, col, attrs, body = int(c.group(2)), _col(c.group(1)), c.group(3), c.group(4) or ""
            sm = re.search(r's="(\d+)"', attrs)
            fill, struck = xs[int(sm.group(1))] if sm and int(sm.group(1)) < len(xs) else ("-", False)
            txt = ""
            if body:
                t = (re.search(r't="(\w+)"', attrs) or [None, "n"])[1]
                v = re.search(r"<v(?: [^>]*)?>(.*?)</v>", body, re.S)
                if t == "s" and v:
                    txt = _live(ss[int(v.group(1))], struck)
                elif t == "inlineStr":
                    txt = _live((_txt(body), _runs(body)), struck)
                elif v and not struck:
                    txt = html.unescape(v.group(1))
            if (r, col) in nonanchor:
                txt = ""
            txt = txt.replace("\t", " ").replace("\r", "").replace("\n", "⏎").strip()
            if fill == "-" and not txt:
                continue
            w(f"[{r},{col}]\t{fill}\t{merges.get((r, col), '')}\t{txt}\n")
    if not hit:
        print(f"WARNING: no sheet matched {sys.argv[2]!r}", file=sys.stderr)


if __name__ == "__main__":
    main()

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


def _shared(z):
    try:
        x = z.read("xl/sharedStrings.xml").decode("utf-8")
    except KeyError:
        return []
    out = []
    for si in re.findall(r"<si>(.*?)</si>", x, re.S):
        runs = re.findall(r"<r>(.*?)</r>", si, re.S)
        out.append("".join(_txt(r) for r in runs if not STRIKE.search(r)) if runs else _txt(si))
    return out


def _styles(z):
    x = z.read("xl/styles.xml").decode("utf-8")
    fonts = re.search(r"<fonts[^>]*>(.*?)</fonts>", x, re.S).group(1)
    fstrike = [bool(STRIKE.search(f)) for f in re.findall(r"<font(?:/>|>.*?</font>)", fonts, re.S)]
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
        for mm in re.finditer(r'<mergeCell ref="([A-Z]+)(\d+):([A-Z]+\d+)"', x):
            merges[(int(mm.group(2)), _col(mm.group(1)))] = f"{mm.group(1)}{mm.group(2)}:{mm.group(3)}"
        w(f"## {name}\n")
        for c in re.finditer(r'<c r="([A-Z]+)(\d+)"([^>]*?)(?:/>|>(.*?)</c>)', x, re.S):
            r, col, attrs, body = int(c.group(2)), _col(c.group(1)), c.group(3), c.group(4) or ""
            sm = re.search(r's="(\d+)"', attrs)
            fill, struck = xs[int(sm.group(1))] if sm and int(sm.group(1)) < len(xs) else ("-", False)
            txt = ""
            if body and not struck:
                t = (re.search(r't="(\w+)"', attrs) or [None, "n"])[1]
                v = re.search(r"<v>(.*?)</v>", body, re.S)
                if t == "s" and v:
                    txt = ss[int(v.group(1))]
                elif t == "inlineStr":
                    txt = _txt(body)
                elif v:
                    txt = html.unescape(v.group(1))
            txt = txt.replace("\t", " ").replace("\r", "").replace("\n", "⏎").strip()
            if fill == "-" and not txt:
                continue
            w(f"[{r},{col}]\t{fill}\t{merges.get((r, col), '')}\t{txt}\n")
    if not hit:
        print(f"WARNING: no sheet matched {sys.argv[2]!r}", file=sys.stderr)


if __name__ == "__main__":
    main()

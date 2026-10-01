"""Read a shared-dump .txt correctly: tokens grouped by the [row,col] inside them, never by physical line.

A cell value can contain a newline, so one sheet row can span several physical lines of the dump;
reading line by line silently drops every token after the first embedded newline (this produced a
false "missing join condition" finding on PSJCO501 TSJCA311 row 70). Values are returned verbatim —
never trim them; leading/trailing spaces are significant in these workbooks.

    import sys; sys.path.insert(0, r"<plugin>/skills/_shared/scripts")
    from dump_cells import load, rows
    cells = load(path)          # {(row, col): value}
    R = rows(cells)             # {row: {col: value}}

CLI:  python dump_cells.py <dump.txt> [row [row2]]   -> prints those rows as "row: col=value | ..."
"""
import collections, re, sys

_TOK = re.compile(r"\[(\d+),(\d+)\]=")

def load(path):
    text = open(path, encoding="utf-8-sig").read().replace("\r\n", "\n")
    ms = list(_TOK.finditer(text)); cells = {}
    for i, m in enumerate(ms):
        end = ms[i + 1].start() if i + 1 < len(ms) else len(text)
        v = text[m.end():end]
        if v.endswith(" | "): v = v[:-3]          # separator before the next token on the same line
        elif v.endswith("\n"): v = v[:-1]         # end of the physical line that closes this row
        cells[(int(m.group(1)), int(m.group(2)))] = v
    return cells

def rows(cells):
    r = collections.defaultdict(dict)
    for (a, b), v in cells.items(): r[a][b] = v
    return dict(r)

if __name__ == "__main__":
    if len(sys.argv) < 2: sys.exit(__doc__)
    R = rows(load(sys.argv[1]))
    lo = int(sys.argv[2]) if len(sys.argv) > 2 else min(R)
    hi = int(sys.argv[3]) if len(sys.argv) > 3 else (lo if len(sys.argv) > 2 else max(R))
    for a in sorted(R):
        if lo <= a <= hi:
            print(f"{a}: " + " | ".join(f"{b}={R[a][b]!r}" for b in sorted(R[a])))

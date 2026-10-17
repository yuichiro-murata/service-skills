"""Zoom list for design-doc-writing-rules W7: every live row of the visible `ｽﾞｰﾑ一覧*` sheets of the zoom ledgers
directly in 01_Doc\\04_共通設計 (03.ｽﾞｰﾑ管理台帳.xlsx = 共通, 03.ｽﾞｰﾑ管理台帳_<WG>.xlsx = 個別; never
90_JAGUR各管理台帳\\, never hidden sheets such as 旧ｽﾞｰﾑ管理台帳(共通)). Reads the zip only (xlsx_zip.py), so the
1,048,566-row used range of _基準情報's ｽﾞｰﾑ一覧【XJ】 costs nothing; struck/gray rows are dropped. The header row
is found by its `ｽﾞｰﾑID` label and every column by label. Duplicate IDs (a changed row repeated below the original,
the same ID in several ledgers) are all kept — accept any live name of an ID.

Output TSV: zoom_id, name, display_name, kobetsu_kyotsu, owner, ledger, sheet, row — plus a `# source:` line per
ledger (mtime-utc, length) that --check compares as values.

usage: python build_zoom_list.py <...\\01_Doc\\04_共通設計> <out.tsv> [--check]
"""
import glob, os, re, sys

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from xlsx_zip import iter_sheets  # noqa: E402

LIST = re.compile(r"^ｽﾞｰﾑ一覧")
ZID = re.compile(r"^Z[A-Z]{3}\d{3}$")
COLS = {"ｽﾞｰﾑID": "id", "ｽﾞｰﾑ名称": "name", "表示名": "disp", "個/共": "kk", "主管": "owner"}


def ns(s):
    return re.sub(r"\s", "", s or "")


def ledgers(root):
    return sorted(p for p in glob.glob(os.path.join(root, "03.ｽﾞｰﾑ管理台帳*.xlsx")) if not os.path.basename(p).startswith("~$"))


def sources(files):
    return [f"# source: {os.path.basename(p)}\t{int(os.path.getmtime(p))}\t{os.path.getsize(p)}" for p in files]


def main():
    args = [a for a in sys.argv[1:] if a != "--check"]
    if len(args) != 2:
        sys.exit(__doc__)
    root, out = args
    files = ledgers(root)
    if not files:
        sys.exit(f"no 03.ｽﾞｰﾑ管理台帳*.xlsx directly under {root}")
    sys.stdout.reconfigure(encoding="utf-8")
    if "--check" in sys.argv:
        try:
            old = [l.rstrip("\n") for l in open(out, encoding="utf-8") if l.startswith("# source:")]
        except OSError:
            old = []
        print("FRESH" if old == sources(files) else "STALE")
        return
    rows = []
    for p in files:
        for sheet, _st, cells in iter_sheets(p, name_filter=LIST, max_col=60):
            hdr = cmap = None
            for r in sorted(cells):
                lab = {ns(v): c for c, v in cells[r].items()}
                if hdr is None:
                    if "ｽﾞｰﾑID" in lab:
                        hdr = r
                        cmap = {k: lab[l] for l, k in COLS.items() if l in lab}
                    continue
                zid = ns(cells[r].get(cmap["id"], ""))
                if not ZID.match(zid):
                    continue
                g = lambda k: (cells[r].get(cmap[k], "") if k in cmap else "").replace("\t", " ").replace("\n", "⏎").strip()
                rows.append((zid, g("name"), g("disp"), g("kk"), g("owner"), os.path.basename(p), sheet, str(r)))
    with open(out, "w", encoding="utf-8") as f:
        f.write("# lookup: zoom-id -> ｽﾞｰﾑ名称 (03.ｽﾞｰﾑ管理台帳*, visible ｽﾞｰﾑ一覧* sheets, live rows)\n")
        f.write("\n".join(sources(files)) + "\n")
        f.write("# columns: zoom_id<TAB>name<TAB>display_name<TAB>kobetsu_kyotsu<TAB>owner<TAB>ledger<TAB>sheet<TAB>row\n")
        f.writelines("\t".join(r) + "\n" for r in rows)
    print(f"ledgers={len(files)} rows={len(rows)} ids={len({r[0] for r in rows})}")


if __name__ == "__main__":
    main()

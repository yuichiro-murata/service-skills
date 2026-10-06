"""入力可文字種 reference table for design-doc-writing-rules W1f (self-check No.42): one line per
TextBox/TextArea row of 画面設計書 Ⅴ．画面項目定義 across a WG's design-doc folder.

Scope: if <root> has PHASE* sub-folders, only the workbooks directly inside them (the 工程管理 layout;
their sub-folders hold checklists and old copies); otherwise <root> recursively. Stale copies are skipped
everywhere: folders bk / draw.io / 高S<n>対応 / 90_Branches / 開発DDL作成用 / 90_JAGUR各管理台帳, files
完了_* and the XXXXX000 sample, and sheets other than `画面設計書(<ID>)`. The program ID is taken from
the file name by pattern (`完了_阿部_PXJDO101_…` is PXJDO101), so a program's own copy is never an
"other program". Columns are found by header label (画面項目名 / 属性 /
入力可文字種 or 入力可 / 画面項目ID), whitespace-normalised. Cells are read live with `live_dump.classify` (read_only +
rich_text): struck/gray text is dropped and a partly-struck cell keeps its live part — `PXJCO101`
`GXJC101A` `[620,31]` is struck `-` + live `文字列(半英数)`, which a plain read joins into `-文字列(半英数)`. Rows whose 文字種
is `-`, blank or a `※n` footnote are skipped — W1e covers the missing ones.

Output TSV: program, sheet, row, col (of the 文字種 cell), name, norm_name, kind, id. Freshness: `# source-count` and
`# source-newest-mtime-utc` header lines (compare as values).

usage: python build_kind_table.py <WG folder, e.g. ...\\01_Doc\\08_機能定義書\\11_工程管理> <out.tsv> [--check]
       --check: print FRESH/STALE for an existing <out.tsv> and exit 0/1 without rebuilding.
"""
import datetime, glob, os, re, sys, warnings, multiprocessing as mp
import openpyxl
warnings.filterwarnings("ignore")
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from live_dump import classify

STALE_DIR = re.compile(r"\\(bk|draw\.io|高S\d+対応|90_Branches|開発DDL作成用|90_JAGUR各管理台帳|jira_slack_notifier)\\", re.I)
STALE_FILE = re.compile(r"^(~\$|完了_)|XXXXX000")
SHEET = re.compile(r"^画面設計書\([^)]+\)\s*$")
PROG = re.compile(r"[PS][XS]J[A-Z][OB][0-9A-Z]{3}")

FORMAT_VERSION = "2"

def check_fresh(out, files):
    """--check: FRESH only when format-version, source-count and newest mtime (as a value) all match."""
    try:
        head = {}
        for line in open(out, encoding="utf-8-sig"):
            if not line.startswith("#"): break
            if ":" in line: k, v = line[1:].split(":", 1); head[k.strip()] = v.strip()
        newest = max(os.path.getmtime(p) for p in files)
        ok = (head.get("format-version") == FORMAT_VERSION and head.get("source-count") == str(len(files))
              and abs(datetime.datetime.fromisoformat(head["source-newest-mtime-utc"]).timestamp() - newest) < 1)
    except Exception:
        ok = False
    print("FRESH" if ok else "STALE"); sys.exit(0 if ok else 1)
NK = lambda s: re.sub(r"\s+", "", str(s or ""))

def norm_name(s):
    s = NK(s)
    s = re.sub(r"\((FROM|TO|From|To)\)$", "", s)
    s = re.sub(r"\(\d+(-\d+)?\)$", "", s)
    return re.sub(r"[0-9０-９①-⑳]+$", "", s)

def scan(p):
    out = []
    try:
        wb = openpyxl.load_workbook(p, read_only=True, data_only=True, rich_text=True)
    except Exception as e:
        return [("ERR", os.path.basename(p), str(e)[:80])]
    m = PROG.search(os.path.basename(p))
    prog = m.group(0) if m else os.path.basename(p).split("_")[0]
    for ws in wb.worksheets:
        if not SHEET.match(ws.title) or ws.sheet_state != "visible": continue
        inV = False; cols = None
        for row in ws.iter_rows(max_col=60):
            vals = {}
            for c in row:
                if getattr(c, "value", None) is None: continue
                live = classify(c)[0]
                if live is not None and live.strip(): vals[c.column] = live
            if not vals: continue
            r = row[0].row if hasattr(row[0], "row") else None
            first = vals.get(2, "")
            if first.startswith("Ⅴ"): inV = True; continue
            if inV and first.startswith("Ⅵ"): break
            if not inV: continue
            lab = {NK(v): k for k, v in vals.items()}
            if "画面項目名" in lab and ("入力可文字種" in lab or "入力可" in lab):
                cols = (lab["画面項目名"], lab.get("属性"), lab.get("入力可文字種", lab.get("入力可")), lab.get("画面項目ID"))
                continue
            if cols is None or not re.fullmatch(r"\d+", NK(vals.get(3, ""))): continue
            name, attr, kind, iid = (NK(vals.get(k, "")) if k else "" for k in cols)
            if attr not in ("TextBox", "TextArea") or kind in ("", "-") or kind.startswith("※"): continue
            out.append((prog, ws.title, str(r), str(cols[2]), name, norm_name(name), kind, iid))
    wb.close()
    return out

def main():
    args = [a for a in sys.argv[1:] if a != "--check"]
    if len(args) != 2: sys.exit(__doc__)
    root, out = args
    phases = [d for d in glob.glob(os.path.join(root, "PHASE*")) if os.path.isdir(d)]
    files = []
    if phases:   # only the workbooks directly in PHASE*: sub-folders (draw.io\チェックシート…, JAGUR版) hold old copies
        for d in phases:
            files += [p for p in glob.glob(os.path.join(d, "*.xls[xm]")) if not STALE_FILE.search(os.path.basename(p))]
    else:
        for dp, _dn, fn in os.walk(root):
            for f in fn:
                p = os.path.join(dp, f)
                if f.lower().endswith((".xlsx", ".xlsm")) and not STALE_FILE.search(f) and not STALE_DIR.search(p):
                    files.append(p)
    if not files: sys.exit(f"no workbooks found under {root}")
    if "--check" in sys.argv: check_fresh(out, files)
    with mp.Pool(min(8, os.cpu_count() or 4)) as pool:
        res = [x for r in pool.map(scan, sorted(files)) for x in r]
    errs = [x for x in res if x[0] == "ERR"]; rows = [x for x in res if x[0] != "ERR"]
    newest = max(os.path.getmtime(p) for p in files)
    with open(out, "w", encoding="utf-8") as f:
        f.write("# lookup: 入力可文字種 per TextBox/TextArea row of 画面設計書 Ⅴ\n")
        f.write(f"# scope: {root} ({'PHASE* only' if phases else 'recursive'})\n")
        f.write(f"# format-version: {FORMAT_VERSION}\n")
        f.write(f"# source-count: {len(files)}\n")
        f.write(f"# source-newest-mtime-utc: {datetime.datetime.fromtimestamp(newest, datetime.timezone.utc).isoformat()}\n")
        f.write("# columns: program<TAB>sheet<TAB>row<TAB>col<TAB>name<TAB>norm_name<TAB>kind<TAB>id\n")
        for r in rows: f.write("\t".join(x.replace("\t", " ") for x in r) + "\n")
    print(f"files={len(files)} rows={len(rows)} errors={len(errs)}")
    for e in errs: print("ERR", e[1], e[2])

if __name__ == "__main__":
    main()

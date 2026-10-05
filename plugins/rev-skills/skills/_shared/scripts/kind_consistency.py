"""W1f (self-check No.42) candidates for one program from build_kind_table.py's TSV.

(a) corpus: a row whose 入力可文字種 differs from the value other programs use for the same item
    (normalised 画面項目名), when that value is dominant — at least MIN_N rows from at least MIN_PROG
    distinct other programs, and a share of at least SHARE. Mixed items (KC品名, 処置指示No …) have no
    dominant value and are skipped. Rows of one item with the same value are grouped on one line.
(b) program: the same item on two of the program's own screens with different 文字種 (always shown;
    the agent judges, e.g. 年月日 vs 年月 may be intended).
Items are matched by name, never by 画面項目ID alone: a wrong ID (作業工程GRP carrying 製造ﾛｯﾄNo's
XJC0028) is design-doc-internal-consistency check 3's finding, not a 文字種 one.
Candidates only — confirm each row in the live dump before reporting.

usage: python kind_consistency.py <kind_table.tsv> <program id, e.g. PSJCO204> [--min-n 5] [--min-prog 3] [--share 0.8]
"""
import argparse, collections

def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("table"); ap.add_argument("program")
    ap.add_argument("--min-n", type=int, default=5); ap.add_argument("--min-prog", type=int, default=3)
    ap.add_argument("--share", type=float, default=0.8)
    a = ap.parse_args()
    rows = []
    a.program = a.program.upper()
    for line in open(a.table, encoding="utf-8-sig"):
        if line.startswith("#") or not line.strip(): continue
        prog, sheet, r, c, name, nname, kind, iid = line.rstrip("\n").split("\t")
        rows.append(dict(prog=prog, sheet=sheet, r=r, c=c, name=name, nname=nname, kind=kind, id=iid))
    by = collections.defaultdict(list)
    for x in rows: by[x["nname"]].append(x)
    mine = [x for x in rows if x["prog"] == a.program]
    if not mine:
        print(f"no TextBox/TextArea rows for {a.program} in the table"); return
    print("# (a) differs from the dominant value in other programs (one line per item and value)")
    hits = collections.defaultdict(list)
    for x in mine:
        others = [y for y in by[x["nname"]] if y["prog"] != a.program]
        if len(others) < a.min_n: continue
        c = collections.Counter(y["kind"] for y in others); dom, dn = c.most_common(1)[0]
        progs = {y["prog"] for y in others if y["kind"] == dom}
        if dn / len(others) >= a.share and len(progs) >= a.min_prog and x["kind"] != dom:
            hits[(x["nname"], x["kind"], dom, dn, len(others), len(progs))].append(x)
    for (nn, kind, dom, dn, n, np_), xs in hits.items():
        cells = ", ".join(f"{x['sheet']} [{x['r']},{x['c']}]" for x in xs)
        print(f"{nn}\t{kind}\tvs {dom} ({dn}/{n} rows, {np_} programs)\t{cells}")
    print("# (b) same item, different 文字種 within this program")
    g = collections.defaultdict(list)
    for x in mine: g[x["nname"]].append(x)
    for k, v in g.items():
        if len({y["kind"] for y in v}) > 1:
            print(f"{k}\t" + " | ".join(f"{y['sheet']} [{y['r']},{y['c']}] {y['kind']}" for y in v))

if __name__ == "__main__":
    main()

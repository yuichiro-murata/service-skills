"""区分名称 index for design-doc-internal-consistency check 11: {group: {区分: 区分名称}} from the live
STEP2 sheet of 01_Doc\\04_共通設計\\09.区分名称_step2.xlsx.

Structure (measured): a group title in col 4, then a header row 区分(col 5) / 区分名称(col 10), then
value rows; groups are separated by a blank row. A title that follows its predecessor with NO blank
row is a sub-heading (`合否判定条件` → `入力ﾀｲﾌﾟ 1:文字の場合` …): its values merge into the parent
group, and a code that means different things under different sub-headings keeps all names joined
by " | ". A title naming several groups (`ﾛｯﾄｶｰﾄﾞ発行区分、ﾏｶﾞｼﾞﾝｶｰﾄﾞ発行区分`) is split on 、.
Struck/gray content is excluded (the sheet is read through live_dump.py).

Freshness: the JSON carries source-mtime-utc and source-length; compare them as a datetime / int
with the file (not as strings — Python and PowerShell print different fractional digits).

usage: python build_kbn_index.py <09.区分名称_step2.xlsx> <out.json>
"""
import datetime, json, os, re, shutil, subprocess, sys, tempfile
HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, HERE)
from dump_cells import load, rows

def main():
    if len(sys.argv) != 3: sys.exit(__doc__)
    src, out = sys.argv[1], sys.argv[2]
    tmp = tempfile.mkdtemp(prefix="kbn_")
    try:
        build(src, out, tmp)
    finally:
        shutil.rmtree(tmp, ignore_errors=True)

def build(src, out, tmp):
    subprocess.run([sys.executable, os.path.join(HERE, "live_dump.py"), src, tmp,
                    "--sheets", "^区分名称_STEP2", "--prefix", "KBN", "--max-col", "20"],
                   check=True, stdout=subprocess.DEVNULL)
    txt = [f for f in os.listdir(tmp) if f.startswith("KBN_") and f.endswith(".txt")]
    if len(txt) != 1: sys.exit(f"expected one STEP2 sheet, got {txt}")
    R = rows(load(os.path.join(tmp, txt[0])))
    # rows whose content was struck out are gaps in the live dump but NOT blank separators
    # (09.区分名称 rows 2011-2012 under 製造条件表示区分 are struck; the sub-headings below stay in it)
    dead = {int(m.group(1)) for m in re.finditer(r"\[(\d+),\d+\]",
            open(os.path.join(tmp, "_DELETED_DIGEST_KBN.txt"), encoding="utf-8").read())}
    is_hdr = lambda row: row.get(5, "").strip() == "区分" and row.get(10, "").strip() == "区分名称"
    groups, active, prev = {}, [], None
    for r in sorted(R):
        row = R[r]
        title = row.get(4, "").strip()
        if title and not is_hdr(row) and 5 not in row:
            if prev is None or not active or any(x not in dead for x in range(prev + 1, r)):   # a truly blank row: new group
                active = [x.strip() for x in title.split("、") if x.strip()]
                for g in active: groups.setdefault(g, {})
            # else: sub-heading — keep filling the current group
        elif is_hdr(row):
            pass
        elif active and 5 in row and 10 in row:
            k, v = row[5].strip(), row[10].strip()
            for g in active:
                cur = groups[g].get(k)
                if cur is None: groups[g][k] = v
                elif v not in cur.split(" | "): groups[g][k] = cur + " | " + v
        prev = r
    st = os.stat(src)
    obj = {"source": os.path.basename(src),
           "source-mtime-utc": datetime.datetime.fromtimestamp(st.st_mtime, datetime.timezone.utc).isoformat(),
           "source-length": st.st_size, "groups": groups}
    with open(out, "w", encoding="utf-8") as f: json.dump(obj, f, ensure_ascii=False, indent=0)
    print(f"groups={len(groups)}")

if __name__ == "__main__":
    main()

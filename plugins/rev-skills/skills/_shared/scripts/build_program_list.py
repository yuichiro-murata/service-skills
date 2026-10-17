"""Program-name list for design-doc-writing-rules W8: <機能ID> -> <機能名> from the 機能定義書 FILE NAMES under
01_Doc\\08_機能定義書 (all WGs), in build_screen_list.py's scope (PHASE* top level where a WG has PHASE folders,
else recursive; stale folders/files skipped). Opens no workbook (~1 s), so it is rebuilt every run — no cache.
04_品質管理's branch numbers stay in the key (PXJDO101_010_規格登録 -> PXJDO101_010). Prints DUP-ID / UNPARSED
lines (not findings).

usage: python build_program_list.py <...\\01_Doc\\08_機能定義書> <out.tsv>
"""
import os, re, sys

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from build_screen_list import workbooks  # noqa: E402

ID = re.compile(r"^([A-Z]{3,5}[A-Z0-9]{3,5})(?:_(\d{3}))?_(.+)$")


def main():
    if len(sys.argv) != 3:
        sys.exit(__doc__)
    root, out = sys.argv[1:3]
    sys.stdout.reconfigure(encoding="utf-8")
    rows, seen = [], {}
    for p in sorted(workbooks(root)):
        m = ID.match(os.path.splitext(os.path.basename(p))[0])
        rel = os.path.relpath(p, root)
        if not m:
            print("UNPARSED\t" + rel)
            continue
        pid = m.group(1) + ("_" + m.group(2) if m.group(2) else "")
        if pid in seen:
            print(f"DUP-ID\t{pid}\t{seen[pid]}\t{rel}")
        seen.setdefault(pid, rel)
        rows.append((pid, m.group(3), rel))
    with open(out, "w", encoding="utf-8") as f:
        f.write("# lookup: program-id -> program-name (機能定義書 file name)\n# scope: as build_screen_list.py\n"
                "# columns: program_id<TAB>name<TAB>workbook\n")
        f.writelines("\t".join(r) + "\n" for r in rows)
    print(f"programs={len(rows)}")


if __name__ == "__main__":
    main()

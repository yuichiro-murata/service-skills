# Detecting font-size, font-name and cell-merge irregularities

This is specific to `design-doc-formatting-consistency` — no other REV skill needs this technique,
so it lives in its own file rather than in `_shared/xlsx-excel-com-dump.md` (which every REV skill
reads) to avoid making every other skill pay the token cost of a section it never uses.

A separate class of defect from strikethrough: a cell whose **font size**, **font name** or **merge
span** breaks from the surrounding pattern, usually left over from a copy-paste edit that didn't fully
take (a pasted row keeping the source's font, a header block duplicated on the sheet's right side that
never got updated when the left side did).

**This is the one formatting signal the shared dump does *not* pre-resolve.** Strikethrough and
gray-out are applied while the workbook is open and never reach the agents (see
`xlsx-excel-com-dump.md`), but font and merge span have no equivalent — this skill reads them from the
workbook itself. Since only this skill reads the result, that is correct: don't try to fold it into
the shared dump for the other eight, which never look at it.

## Which cells to scan

Iterate the **live dump's coordinates** (`dump_cells.load()` on the per-sheet `.txt`), not the raw
sheet: that keeps struck/gray cells out without a strike scan of your own (agent-guide forbids one).
Skip non-anchor cells of merged ranges. An older COM dump emitted values Excel had hidden under a
merge, which produced phantom 9pt ＭＳ Ｐゴシック "deviations" at 画面設計書(GXJC125A) `[415,9]`/`[417,9]`
on PXJCO125; the COM template now drops them, and the anchor-only rule guards against any recurrence.

## Routes

**openpyxl (validated end-to-end on PXJCO125; needs no Excel process — prefer it).**
`wb = openpyxl.load_workbook(path, data_only=True, rich_text=True)` — **not** `read_only=True`, which
has no `merged_cells`. Per cell: `cell.font.sz`, `cell.font.name`. A `CellRichText` value with runs of
different sizes is "mixed": read each `TextBlock.font.sz` (run font name is `TextBlock.font.rFont`; a
bare `str` part inherits the cell font) and treat the cell as "mixed, inspect manually". Merges:
`ws.merged_cells.ranges` gives anchor (`min_row`/`min_col`) and span (`max_row-min_row+1`,
`max_col-min_col+1`). Coordinates are sheet-absolute, so no UsedRange offset applies. Cost: the load
alone took **124 s** for PXJCO125 (45 sheets, 41 visible), with a harmless "wmf image format is not
supported" warning. If openpyxl cannot open the workbook (agent-guide: `PXJCO130` raises
`xl/drawings/NULL`), use COM.

**COM.** First read, in `xlsx-excel-com-dump.md`, `## The dump script` (the pre-existing-PID guard
that stops you `Quit()`-ing the user's own Excel, and the Add-Type/inline rules),
`## UsedRange-relative vs sheet-absolute coordinates` and `## Operational notes`. Read `Font.Size` /
`Font.Name` per cell with the DBNull guard for mixed-formatting runs (`if ($cell.Font.Size -is
[System.DBNull]) { ... }` — "mixed, inspect manually", not a crash). Merges: `cell.MergeCells`,
anchor when `cell.MergeArea.Row -eq r -and .Column -eq c`, span `MergeArea.Rows.Count`/`.Columns.Count`.
**Apply the UsedRange origin offset**: bulk `Value2` coordinates are UsedRange-relative while
`Cells.Item(r,c)` is absolute, so add `$used.Row - 1` / `$used.Column - 1` and report in the absolute
frame the dump uses (12 of the 600 in-scope PHASE3 sheets have a non-A1 origin).

## Font size and font name

Per sheet, tally size and name frequencies and take each mode; cells that differ are candidates.
Font name matters as much as size: on PXJCO125, 更新条件表(TXJCA003) `[70,4]` '投入ﾛｯﾄNo' was
ＭＳ Ｐゴシック 9pt in an item list that is otherwise ＭＳ ゴシック 10pt — the only real font defect in
that sheet, and the name was the clearer signal.

**Most of what this turns up is a deliberate, consistent convention, not a defect.** Silently exclude
these recurring patterns:
- **The revision-memo column**, typically column 105/106 but also 107, 119, 123, 125 or 126 (PXJCO125:
  機能定義書 `[81,123]`, ﾁｪｯｸ処理設計書 `[35,123]`/`[36,123]`, TSJCD014 `[45,126]`). Detect it by content —
  dated `yyyy/m/d 名前 修正`-style entries right of the print area (column ≥ 104) — not by column
  number. It is often smaller and sometimes ＭＳ Ｐゴシック.
- The sheet-title/section-header template cells every sheet shares (row 1's `2C`/`3C` markers, the
  sheet-name label, `Ⅰ．`–`Ⅷ．` section headings) — a fixed, larger size across all workbooks.
- **表紙's sections at 12pt.** 表紙's mode is 10pt, but its section headings and the Ⅰ．上位文書 /
  Ⅱ．設計書構成 / Ⅲ．改訂履歴 table headers and No./name columns are 12pt by template (plus the 24pt
  title and 28pt marker) — 113 "deviations" on PXJCO125, none real. Exempt 12pt on 表紙; compare 表紙
  cells only within their own section.
- Column-group headers in a `Ⅵ．画面項目制御`-style matrix (e.g. `処理区分="..."の場合` labels) —
  deliberately smaller, consistently so across every instance of that matrix.
- **Two-line matrix cells shrunk to fit** — e.g. 画面設計書(GXJC125A) `[544,13]`/`[544,23]`/`[544,28]`
  are 6pt holding `○⏎※1` in a one-row control-matrix cell.
- Any 画面イメージ/mockup-screenshot sheet — a UI mockup naturally mixes sizes; never flag it.

Only report a deviation that's **genuinely isolated** — a single cell or a small handful that fit none
of the patterns above, especially inside an otherwise-uniform repeating list (PXJCO125 機能定義書
`[81,23]`/`[81,24]`: the ○ marks of one Ⅲ．入出力定義 row at 9pt where every other ○ is 10pt).

## Cell merges

**Do not compare merge spans by row proximity or sheet-wide by column position.** Both were tried.
Sheet-wide comparison surfaced tens of thousands of false positives. The "same column, spans differ,
within ~4 rows" rule gave **891 candidates on PXJCO125, almost none real**: most were row-height
differences between multi-row records (表紙 改訂履歴 No. cells 2×2 / 3×2 / 4×2 / 18×2), the rest a
label row beside a value row (更新ﾃｰﾌﾞﾙ 1×14 vs 更新概要 3×14; 検索条件 1×42 vs `No.` 1×2) repeated on
every 更新条件表. Instead:

1. **Compare column spans only.** Height differences between records of a list are expected (a
   record grows with its text).
2. **Compare only cells of the same role** — the same normalised label (whitespace stripped) in the
   same column, or sibling records of one list (same column, between the list's header row and its
   footer such as `特記事項`). Never pair a label cell with a value cell.
3. **Subtract template signatures** — `(column, span A, span B, normalised labels)` tuples that recur
   on 3 or more sheets of the workbook, or across sibling workbooks in the same PHASE folder, are the
   template, not a defect.

What survives is a lead for the content cross-check (SKILL.md step 3): a duplicated header block whose
right-side copy kept stale content, or one list row merged differently from its siblings. A whole
block shifted off the template's columns (PXJCO125 更新条件表(TXJCM006): its 取得条件 blocks sit one
column right of every other 更新条件表) is a different shape: catch it by comparing a sheet's anchor
columns against the sibling sheets of the same template, not by span.

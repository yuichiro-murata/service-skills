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
the shared dump for the other nine, which never look at it.

## Which cells to scan

Iterate the **live dump's coordinates** (`dump_cells.load()` on the per-sheet `.txt`), not the raw
sheet: that keeps struck/gray cells out without a strike scan of your own (agent-guide forbids one).
Skip non-anchor cells of merged ranges. An older COM dump emitted values Excel had hidden under a
merge, which produced phantom 9pt ＭＳ Ｐゴシック "deviations" at 画面設計書(GXJC125A) `[415,9]`/`[417,9]`
on PXJCO125; the COM template now drops them, and the anchor-only rule guards against any recurrence.

**Scan only the template sheets** — those named `表紙`, `機能定義書`, `画面設計書`, `ﾁｪｯｸ処理設計書`,
`更新条件表`, `ﾌｧｲﾙ…仕様書`, `帳票設計書` (`詳細設計書` per SKILL.md). Mock-up and reference sheets are
rarely named 画面ｲﾒｰｼﾞ: `PXJCO161` (2026-10-06) has `実績入力_B画面`…`_I画面` (all Meiryo UI),
`(参考)…`, `【JAGUR】(参考)…` and `計算ﾎﾞﾀﾝ押下時の計算方法` (ＭＳ Ｐゴシック, no `Ⅰ．` header). Designer scratch
sheets (`PSJCO403` `資料`, 401 raw deviations) are the same case. Exclude them all and list them once in
one line ("書式ﾁｪｯｸ対象外: …"); never report their cells.

## Routes

**openpyxl (validated end-to-end on PXJCO125; needs no Excel process — prefer it).**
`wb = openpyxl.load_workbook(path, data_only=True, rich_text=True)` — **not** `read_only=True`, which
has no `merged_cells`. Per cell: `cell.font.sz`, `cell.font.name`. A `CellRichText` value with runs of
different sizes is "mixed": read each `TextBlock.font.sz` (run font name is `TextBlock.font.rFont`; a
bare `str` part inherits the cell font) and treat the cell as "mixed, inspect manually". Merges:
`ws.merged_cells.ranges` gives anchor (`min_row`/`min_col`) and span (`max_row-min_row+1`,
`max_col-min_col+1`). Coordinates are sheet-absolute, so no UsedRange offset applies. Cost: the load
alone took **124 s** for PXJCO125 (45 sheets, 41 visible), with a harmless "wmf image format is not
supported" warning, and **~540 s for the 9 MB `PXJCO161`** — over the default tool timeout. Run the
extraction in the background (or at the maximum timeout), load once, and pickle what you need — per
dump coordinate `(value, size, name, runs)` plus every sheet's full merge-range list — then analyse
from the pickle. If openpyxl cannot open the workbook (agent-guide: `PXJCO130` raises
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
  sheet-name label, `Ⅰ．`–`Ⅷ．` section headings) — a fixed, larger size across all workbooks. Also
  row 1's `ｼｽﾃﾑID`/`計画書No` values in the left copy at 11pt (`[1,34]`/`[1,46]`, vs 10pt in the right
  copy `[1,86]`/`[1,98]`) — identical on all four 画面設計書 sheets of `PSJCO403`, i.e. inherited from the
  画面設計書 template copy, not a defect.
- **表紙 has no single mode** — 10pt on PXJCO161 (530 of 662 cells), 12pt on PSJCO403 (80 of 86). Its
  section headings, the Ⅰ．上位文書 / Ⅱ．設計書構成 / Ⅲ．改訂履歴 table headers and the 改訂履歴 No. column
  are 12pt by template (plus the 24pt title and 28pt marker) — 113 "deviations" on PXJCO125, none real.
  On 表紙, compare each cell only with the cells of **its own column within its own section**.
- Column-group headers in a `Ⅵ．画面項目制御`-style matrix (e.g. `処理区分="..."の場合` labels) —
  deliberately smaller, consistently so across every instance of that matrix.
- **Two-line matrix cells shrunk to fit** — e.g. 画面設計書(GXJC125A) `[544,13]`/`[544,23]`/`[544,28]`
  are 6pt holding `○⏎※1` in a one-row control-matrix cell. Exempt only when the siblings with the same
  shape are shrunk the same way; a shrunk cell among unshrunk siblings (or the reverse) is a candidate.

**Compare rich-text runs, not just the cell font.** Within one list, the same token (`※n`) is usually
set in its own run size; compare that run's size across siblings. `PXJCO161` 画面設計書(GXJC161B)
`[2194,22]`/`[2196,22]`/`[2197,22]` hold `n⏎※10` with `※10` at 8pt, but `[2195,22]` is uniformly 10pt
— invisible to a cell-level font scan, caught only at run level.
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
2. **Compare only cells of the same role inside ONE table** — sibling records of one list or block.
   Operationally: a block starts at a row containing `No.`, `取得項目`, `検索条件`, `結合条件(…)` or `ｿｰﾄ順`
   and ends at the next such row (or `取得件数` / a footer such as `特記事項`). Its **data rows** are those
   whose No. cell holds a number; take the per-column span mode over the data rows only and flag the data
   rows that differ (a 2026-10-06 validation run without this rule produced 1,002 raw candidates).
   Never pair a label cell with a value cell,
   and never pair cells across tables because their labels match: Ⅲ's query blocks and Ⅳ/Ⅴ/Ⅵ reuse
   the same names (`画面`, `ﾛｸﾞｲﾝ情報`, `A.ﾎﾞﾃﾞｨ1`) at different spans by design. Skip placeholder cells
   (`-`, blank). The sheet-wide "same normalised label in the same column" rule gave **~230 candidates
   on PSJCO403 画面設計書 A-D, none real**, and the 3-sheet signature exclusion below did not thin them.
3. **Subtract template signatures** — `(column, span A, span B, normalised labels)` tuples that recur
   on 3 or more sheets of the workbook are the template, not a defect. (Do not try to build the
   signature across sibling workbooks on the fly — each 9 MB load is ~9 min; use it only if a prebuilt
   list is handed to you.)
4. **Use every merge range, including ranges over blank cells** — a missing merge usually sits on empty
   cells (`PXJCO161` ﾁｪｯｸ処理設計書(GXJC161B) `[158-160]`: no col-3 1×8 merge where every sibling row has
   one; the label moved to col 7 as 1×4). Filter to dump coordinates only when *reporting*, so a struck
   row is not cited.

What survives is a lead for the content cross-check (SKILL.md step 3): a duplicated header block whose
right-side copy kept stale content, or a few rows inside one block merged differently from the rest
of that block — the useful shape, usually a paste remnant. The procedure that found one: per sheet,
tally the anchor column and span of every operator cell (`=`, `IN`, `LIKE`, `DESC`…). On PSJCO403
画面設計書(GSJC403A) ~300 sit at col 26 as 1×4, but (3-2)'s 検索条件 `[229-233]` put `=` at col 22 as 1×1 (and
the value at col 26 as 1×22) — the one block laid out differently, worth the content check. (A `結合条件` label such as
`[559,26]` is 1×1 everywhere and is not an outlier.)

**Out of scope for the font and merge checks**: undated designer memos and pasted SQL right of the
print area (body rows, cols 53-104 — e.g. GSJC403A `[19,55]`, `[281,54]`). They follow no template, so
every cell would "deviate". Whether such leftovers should stay is one 低 content line at most, not a
per-cell font or merge finding. A whole
block shifted off the template's columns (PXJCO125 更新条件表(TXJCM006): its 取得条件 blocks sit one
column right of every other 更新条件表) is a different shape: catch it by comparing a sheet's anchor
columns against the sibling sheets of the same template, not by span.

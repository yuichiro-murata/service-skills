---
name: design-doc-formatting-consistency
description: Scan a design-doc workbook for stray font-size and cell-merge irregularities — a cell whose font size or merge span breaks from the surrounding pattern, which usually signals a leftover copy-paste artifact, an incompletely-applied edit, or a duplicated header block that never got updated. This is a formatting/editing-hygiene check, distinct from design-doc-internal-consistency (cross-reference completeness), design-doc-io-table-check (I/O table completeness), xlsx-db-column-check (DB column existence), and design-doc-typo-check (actual Japanese-language proofreading). Use as part of a full program REV, or when the user specifically asks about フォントサイズ/セル結合のブレ. For a full REV without named checks the entry point is `rev-program-review`; run this standalone only when selected there or asked for by name.
---

# design-doc-formatting-consistency

Finds stray font-size and cell-merge irregularities across every sheet of one program's design-doc
workbook — see the frontmatter `description` above for what these irregularities typically signal.
Unlike the other REV skills, this one isn't grounded in the project's checklist or a cross-reference
rule; it's a pure structural-hygiene sweep, worth surfacing to the designer even though these aren't
"wrong" in the same sense as a broken cross-reference.

## Environment

**Read `_shared/agent-guide.md` first** — scope selection, environment, reporting conventions — then
`_shared/xlsx-formatting-scan.md` **in full before running anything** (this is the reading agent-guide's
doc map assigns this check). Font size, font name and merge spans are not in the text dump, so this
check reads the workbook's formatting itself; the scan file gives both routes (openpyxl, validated
end-to-end on PXJCO125, and COM), names the `xlsx-excel-com-dump.md` sections to read only if you take
the COM route, and — critically — the noise filters learned from real runs: a naive "most common
size/span" comparison floods the report with template structure. (The technique lives in its own file
because no other REV skill needs it — don't merge it back.)

**Left/right header copies** (`2C`/`3C` vs `2D`/`3D`, rows 1-4): read both copies and compare them
label by label (pair each label cell with its value cell in the left copy, the same label in the right
copy, and compare values). Report a sheet only when the two copies **contradict** — two different
non-`-` values for the same label. A copy reading `-` (or empty) beside a filled one is template-level
drift: on PHASE2, 209 of 424 sheets have differing C/D 更新日 and ~182 sheet pairs show the `-` pattern,
so summarise it as **one 低 line per workbook**, never per sheet (the orchestrator's Step 3.5 rule 2
withdraws per-sheet `-` findings anyway). When `naming-standard-compliance` is in the same run, it owns
both the value contradictions and that one drift line — report only formatting (font/merge)
differences between the copies.

**Don't dump or scan `詳細設計書` sheets — this is a deliberate coverage-for-tokens trade-off, not a
claim that the sheet is always empty.** Unlike `design-doc-internal-consistency`/
`naming-standard-compliance` (which skip it because their specific checks genuinely don't need
detail-design content), this skill's blanket per-cell scan *can* find real things there — a real
run against a sub-program (SXJCB147, a ｻﾌﾞﾌﾟﾛ) found two genuine font-size irregularities inside its
`詳細設計書` sheet (in its ﾊﾟｯｹｰｼﾞ構成 and 引数一覧 sections, which were populated, not
header-only). The user explicitly chose to exclude it anyway for the token savings, accepting that
findings of that kind will be missed going forward. If the user asks for a more thorough pass, or
specifically asks about 詳細設計書 formatting, include it then.

## Procedure

### 1. Font-size and font-name scan

For every **template sheet** of the target workbook (表紙/機能定義書/画面設計書/ﾁｪｯｸ処理設計書/更新条件表/
ﾌｧｲﾙ…仕様書/帳票設計書 — **not `詳細設計書`**, see above; mock-up, `(参考)`, `【JAGUR】` and scratch sheets
such as `資料` are excluded and listed once, per the shared doc), iterate the **live
dump's coordinates** (so struck/gray cells stay out — don't run your own strike scan) and skip
non-anchor cells of merged ranges. Record each cell's font size and font name, and each rich-text
run's size. Per sheet, take the mode of each and list every cell that differs — this raw list will be
large and mostly noise. On a big workbook, run the openpyxl load in the background and pickle the
extract (the shared doc's "Routes").

**Before reporting anything, filter out the deliberate systematic patterns** documented in the
shared doc (revision-memo columns, section-title/header template cells, 表紙 compared per column within
its section, control-matrix group headers, consistently shrunk two-line matrix cells and cells shrunk to fit their own longer text), and leave undated
memos/pasted SQL right of the print area (body cols 53-104) out of both scans. What's left —
genuinely isolated cells, especially ones inside an otherwise-uniform repeating list — is what's
worth reporting.

### 2. Cell-merge scan

Record every merge range's row/column span (blank cells included — a missing merge sits on empty
cells), then compare **column spans only, among the data rows of ONE block** (the shared doc defines
a block: from a `No.`/`取得項目`/`検索条件` row to the next one), skipping `-`/blank placeholders, and
subtract template signatures (span pairs that recur on 3+ sheets of the workbook). Row-height
differences between multi-row records are expected. The shared doc has the exact rule and why the
proximity-only (891 candidates on PXJCO125) and same-label (~230 on PSJCO403) versions failed. What
survives is a lead: a duplicated header block whose right-side copy kept stale content, or a few rows
inside one block merged differently from the rest — on PSJCO403 GSJC403A, (3-2)'s 検索条件 `[229-233]`
(operator at col 22, 1×1, where ~300 others sit at col 26 as 1×4).

### 3. Cross-check anything suspicious against content, not just shape

A merge or font-size anomaly is a **lead**, not a finding by itself — when something survives the
filtering above, look at what's actually written in the flagged cell(s) before reporting. A
duplicated header block with a mismatched merge span is far more actionable once you've also checked
whether its *content* is stale (e.g. an old system name, a blank date, a wrong ID) — that's often
the real, citable defect, with the formatting anomaly serving as how you found it in the first place.

## Reporting

The output goes to the designer who owns the doc, not into your own working notes — report only
what they'd actually need to act on. Write in the user's language.

For each real finding: cite the exact sheet name and cell coordinate(s), state what's different (the
size/span and what the surrounding pattern is), and — when you cross-checked content per step 3 —
say what that content difference actually is, since that's usually the more useful half of the
finding. Group by sheet only when several findings share one.

Omit entirely: the raw per-sheet size/merge dump, a tally of how many cells were checked, and any
finding that turned out to be one of the documented systematic patterns (revision-memo columns,
template headers, 表紙 12pt sections, control-matrix cells) — those are not worth even a one-line
mention once identified as such, since they're expected and consistent by design. If a whole class of
irregularity couldn't be checked (e.g. a sheet's used range was too large to scan in full), say so
briefly — that's a coverage gap worth flagging, not a process detail. Name the non-template sheets
you excluded in one line, so the reader knows they were not scanned.

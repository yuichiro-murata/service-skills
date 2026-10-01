---
name: rev-program-review
description: Entry point for a full REV of ONE program's design-doc workbook (機能定義書/画面設計書/ﾁｪｯｸ処理設計書/更新条件表/帳票設計書/ﾌｧｲﾙ出力仕様書). Instead of unconditionally running every check, it first presents the 9 single-program check skills as a checkbox list (AskUserQuestion, multiSelect) so the user picks which checks to run BEFORE any dumping or analysis starts, then runs only the selected ones as one combined pass and reports their findings as a single merged report. Use whenever the user asks to REV/レビュー a design-doc workbook without naming the specific checks they want — in that case do NOT launch the individual check skills directly. Skip the checkbox prompt only when the user already named the checks, or explicitly asked for 全部/all/フルREV.
---

# rev-program-review

Runs a REV of one program's design-doc workbook with a **user-selected scope**. The 9 checks below
are individually expensive (each dumps and re-reads a large workbook, several also read WG-folder
DB-layout files and the 01_Doc common-design workbooks), so running all of them when the user only
wanted two wastes a lot of time and tokens. Ask first, then run only what was selected.

## The 9 single-program check skills

| # | skill | what it checks |
|---|-------|----------------|
| 1 | `design-doc-internal-consistency` | 内部相互参照（ｲﾍﾞﾝﾄ⇔処理概要、ﾚｽﾎﾟﾝｽ/画面項目ID/ﾁｪｯｸID登録、画面3点一致、排他制御）＋Ⅴ．画面項目定義の記載ﾙｰﾙ（大文字変換、ﾎﾞﾀﾝID 9xxx、選択□、No ZZ9/XJZ0425、TextBoxの桁数・入力可文字種）＋処理/ﾁｪｯｸ順⇔ﾎﾞﾀﾝ順、項番の飛び・重複、区分名称の09.区分名称突合 |
| 2 | `design-doc-io-table-check` | Ⅲ．入出力定義（CRUD一覧）の双方向網羅性 — 最重量・最高収穫のチェック |
| 3 | `xlsx-db-column-check` | 参照ｶﾗﾑがﾃｰﾌﾞﾙﾚｲｱｳﾄに実在するか＋**画面の桁数とﾃｰﾌﾞﾙﾚｲｱｳﾄの桁数の整合**＋ｺｰﾄﾞ/名称の二重保持ｱﾝﾁﾊﾟﾀｰﾝ(名称がLabelの場合のみ指摘)＋検索条件のTXJAM100保存漏れ |
| 4 | `naming-standard-compliance` | 各ID採番規則・設計書記述ﾙｰﾙ準拠 |
| 5 | `update-condition-completeness` | 更新条件表の網羅性（NOT NULL/主キー/共通項目の設定漏れ、登録日時のUPDATE上書き） |
| 6 | `report-design-check` | 帳票設計書（帳票一覧登録、Ⅱ．帳票仕様の記入漏れ、参照先ｴｲﾘｱｽ、取得項目⇔印字項目） |
| 7 | `design-doc-formatting-consistency` | ﾌｫﾝﾄｻｲｽﾞ/ｾﾙ結合のブレ（体裁の衛生） |
| 8 | `design-doc-typo-check` | 日本語の誤字脱字・変換ミス |
| 9 | `file-output-spec-check` | ﾌｧｲﾙ出力仕様書（ﾌｧｲﾙID一貫性、ﾌｧｲﾙ一覧登録、Ⅱ．ﾌｧｲﾙ出力仕様の記入漏れ、共通ﾀﾞｳﾝﾛｰﾄﾞ注釈、画面項目ID番台/登録、ﾍｯﾀﾞｰ⇔明細、明細の参照先、編集方法、ﾌｧｲﾙﾚｲｱｳﾄ一致） |

Checks 5, 6 and 9 are narrower than the rest in one specific sense: each applies only when the
workbook has the sheets it reads. A program with no `更新条件表(*)` sheet (a pure 照会 screen), no
`帳票設計書(*)` sheet, or no visible `ﾌｧｲﾙ出力仕様書(*)` sheet has nothing for that check to do —
still offer it, but if the user selects it and the sheets aren't there, say so in one line in the
merged report rather than running an empty analysis.

WG-folder-scoped comparison skills are **not** part of this menu — bundle those in only when the
user's request is itself folder-scoped, or they explicitly ask for them.

## Step 1 — 対象ワークブックの確定

Resolve which workbook is being reviewed (a program ID, a path, or the IDE selection). If the target
is genuinely ambiguous, ask for it in the same `AskUserQuestion` call as the scope question below —
don't burn a separate round trip.

## Step 2 — チェック項目の選択（チェックボックス）

Call `AskUserQuestion` with **three `multiSelect: true` questions** — the tool allows at most 4
options per question, so the 9 checks are split 4 + 3 + 2. Present them exactly like this (labels in
Japanese, since the reviewers work in Japanese) — keep this grouping and this option order:

- Question 1 — header `整合性`, question「実施するチェックを選択してください（複数選択可）」
  - 「Ⅲ．入出力定義（CRUD）網羅チェック (推奨)」— `design-doc-io-table-check`
  - 「設計書内部の相互参照チェック (推奨)」— `design-doc-internal-consistency`
  - 「DBカラム実在・桁数整合チェック」— `xlsx-db-column-check`
  - 「ID採番・記述ルール準拠チェック」— `naming-standard-compliance`
- Question 2 — header `更新・帳票`, question「更新条件表・帳票・ﾌｧｲﾙ出力のチェックを選択してください（複数選択可）」
  - 「更新条件表の網羅チェック（NOT NULL/主キー/共通項目）」— `update-condition-completeness`
  - 「帳票設計書チェック」— `report-design-check`
  - 「ﾌｧｲﾙ出力仕様書チェック」— `file-output-spec-check`
- Question 3 — header `誤字・体裁`, question「実施する誤字・体裁チェックを選択してください（複数選択可）」
  - 「誤字脱字チェック」— `design-doc-typo-check`
  - 「フォントサイズ・セル結合のブレチェック」— `design-doc-formatting-consistency`

All three questions go in **one** `AskUserQuestion` call, so all 9 checkboxes appear in a single
prompt. When the target workbook has already been identified and you can see it has no
`更新条件表(*)`, no `帳票設計書(*)` or no visible `ﾌｧｲﾙ出力仕様書(*)` sheet, say so in that option's
`description` (「本ﾌﾞｯｸに帳票設計書ｼｰﾄなし」) rather than dropping the option — the absence is itself
information the reviewer may want to question.

Put the skill name in each option's `description` alongside a one-line summary of what it finds, so
the user can tell the options apart without knowing the skill names by heart. The user can select
none in a group — that group's checks are simply skipped.

### 「Other」の扱い

"Other" lets the user type a scope in free text. Two cases, and they mean opposite things:

- **Other with text** — honor exactly what they typed for that group, even if it names a check from
  another group or narrows the scope further ("入出力定義だけ", "画面項目の順序だけ見て").
- **Other selected but left empty** — read it as "この観点は実施しない": skip **every** check in
  that group, list them under 未実施 in the report, and don't ask again. It is the deliberate way to
  opt a whole group out, so treat it exactly like selecting nothing in that group — never as a
  prompt to re-ask, and never as a reason to fall back to running the group's checks.

If **all three** groups end up with no check to run (nothing selected, or empty "Other"), there is
nothing to review: stop, state plainly that no check was run and that the workbook was not dumped,
and do not fall back to running all 9.

### プロンプトを省略してよいケース

Skip Step 2 and go straight to Step 3 when:

- The user already named the checks ("入出力定義と誤字だけ見て", "誤字チェックして") — run exactly those.
- The user explicitly asked for everything ("全部", "フルREV", "全チェック", "all") — run all 9.
- The user invoked one check skill directly by name — that skill runs standalone; this skill isn't involved.
- `AskUserQuestion` is unavailable (non-interactive / batch / subagent context) — fall back to all 9
  and **state in the report** that the full set was run because the scope couldn't be asked.

Re-ask the scope for each **new** REV request; a selection made for one workbook does not carry over
to the next one unless the user says "同じ観点で" or similar.

## Step 3 — 選択されたチェックの実行

1. **Dump the workbook once, up front** — before launching anything — and hand every check the
   resulting scratchpad text files instead of letting each one re-dump the same workbook. See
   `_shared/xlsx-excel-com-dump.md`, section "dump once, share the text" — that file ships **inside
   this plugin** (`<plugin root>/skills/_shared/`), not under `~/.claude/skills/`; glob
   `**/rev-skills/**/skills/_shared/xlsx-excel-com-dump.md` if the path doesn't resolve. (Exception:
   `design-doc-formatting-consistency` still needs live Excel COM access for its font/merge scan;
   only the bulk text dump is shared.)
   `**/rev-skills/**/skills/_shared/xlsx-excel-com-dump.md` typically matches several copies — the
   git-tracked marketplace copy (`.claude/plugins/marketplaces/rev-skills/plugins/rev-skills/...`)
   **and** one or more older snapshots under `.claude/plugins/cache/rev-skills/<version>/...`. Always
   read the **marketplace** copy: the cache lags behind it, and a stale cached copy has already cost
   a run — the `PSJCO309` dump failed on the `Add-Type` CS0675 bitwise-or error that the marketplace
   copy documents a fix for but the cached `1.1.1` copy predates. The cache is keyed on the
   `version` in `.claude-plugin/plugin.json`, so editing a skill without bumping that version leaves
   every cached copy stale **forever** — the loader sees a version it already has and never re-copies.
   Bump the version in the same commit as any skill-content change.
2. **Build the reference-master index in the same pre-launch step, if a selected check needs it** —
   see `_shared/reference-index.md`. Three checks do: `design-doc-internal-consistency` (screen-item
   IDs), `report-design-check` (the `画面項目ID` column on print items) and `file-output-spec-check`
   (the `画面項目ID` column on a ﾌｧｲﾙ出力仕様書's ﾍｯﾀﾞｰ items) — all want the 画面項目辞書 index, so
   build it once when any of them is selected. That doc instructs the *agent* to
   stop rather than build it, so skipping this
   step silently drops those checks' dictionary-registration test. Build **one index per dictionary
   file the program's screen-item IDs route to** — that doc's routing table maps each ID prefix to
   its `82.画面項目辞書_*.xlsx`, and a program using shared `XJZ`/`SJZ` items needs the `_共通` file
   as well as its own WG's. Subset each to the program's IDs, and pass both paths per file in the
   agent's prompt — the subset to read, the full index to grep.
3. Run the selected checks as one combined pass — in parallel background agents when there are
   several. Follow each selected skill's own SKILL.md as the authority for how that check is done;
   this skill only decides *which* checks run.
4. Do **not** post a status update as each agent finishes. Wait until every check in the batch has
   completed, then compose and post **one** merged report.

## Step 3.5 — 報告前の裏取り（orchestrator が自分でやる）

The agents hand back findings; the orchestrating session is what stands behind them. Before writing
the merged report, re-verify every finding **against the original workbook**, in the orchestrator's
own turn. Two failure modes recur, and the reviewer will ask about both — on the `PXJCO124` REV the
user's very next message after the report was 「取り消し線加味して誤指摘ないか確認して」:

1. **Strikethrough / gray-out.** For every cell a finding cites — the cited cell **and the cells
   used as the comparison basis** — re-read the original with openpyxl and classify it
   `LIVE` / `PARTIAL` / `DEAD`. A finding whose evidence turns out to be struck text is withdrawn.
   The shared live dump already removes struck text, so agents working from it are usually safe, but
   an agent that consulted `_DELETED_DIGEST.txt` may have reasoned about dead content, and a
   `PARTIAL` cell's live remainder can differ from what the agent quoted. Also grep the digest for
   each finding's key term: if the term appears **only** as deleted content, the "missing X" finding
   is really "X was deliberately removed". On `PXJCO124` this pass confirmed all 19 findings' cells
   were live, and it *strengthened* two of them — `[184,8]`/`[184,14]` proved to be `PARTIAL` while
   the stale `[184,21]` was fully `LIVE` (an obvious delete-the-name-but-not-the-value slip), and
   `[1054,5]` proved to have been edited recently without adding the events it was missing.
2. **Template boilerplate.** A finding of the shape "this cell/column/row is blank" on a
   shared-template sheet must be measured across the sibling workbooks in the same PHASE folder
   before it is reported. Two `PXJCO124` findings died this way (0 of 28 and 27 of 28 — see
   `naming-standard-compliance`'s 表紙 section for the specifics). One openpyxl pass over the folder
   settles it and takes under a minute.

State the outcome of this pass in the report. "取り消し線を加味した結果、誤指摘は0件" is itself
information the reviewer wants, and it is what lets them trust the rest of the list.

## Step 4 — 報告

One combined report, findings grouped by check, with the selected scope stated at the top so the
reader knows what was and wasn't looked at — e.g.:

```
## REV結果: PXJCO201_処置指示登録
実施チェック: Ⅲ．入出力定義(CRUD)網羅 / 設計書内部の相互参照 / 更新条件表の網羅 / 誤字脱字
未実施: DBカラム実在 / ID採番・記述ルール / 帳票設計書 / フォントサイズ・セル結合
```

Never imply the workbook passed checks that weren't run.

## Step 5 — Excel 指摘一覧への出力

Reviewers routinely ask for the merged report as a workbook ("Excel一覧に出力して") after reading it
in chat, so offer it once at the end of Step 4 rather than making them ask. Name it
`<プログラムID>_REV指摘一覧_<YYYYMMDD>.xlsx`.

**Where to write it: `C:\Users\<user>\Downloads\`, beside the previous rounds' copies — NOT the
PHASE3 design-doc folder.** This file used to say "next to the reviewed workbook" and cite
`01_Doc/08_機能定義書/11_工程管理/PHASE3/` as the established convention. That was wrong and cost a
run a failed `Workbooks.Open`: a project-wide `find` for `*REV指摘一覧*` matches **nothing** under
`01_Doc`, while `Downloads` holds the real series (`PSJCO204_…_20260910.xlsx`,
`PSJCO204_…_20260911.xlsx`, plus the `.csv` the reviewer pastes back into chat). A findings list is
review correspondence, not a deliverable design document, so it does not belong in the design-doc
tree. Locate the most recent copy by searching `Downloads` before writing, and if the series has
moved, follow where the existing files are rather than this path.

**`Downloads` can legitimately come back empty** — measured `2026-09-17`, the machine held no
`*REV指摘一覧*` file anywhere under the user profile, the whole `PSJCO204`/`PSJCO308` series having
since been cleaned up. That is not a reason to go looking under `01_Doc` (they were never there):
fall back to the measured spec below, and say in the report that no prior copy was available to
copy the layout from.

**Open the most recent one and copy its layout** rather than inventing a shape, and rather than
trusting the numbers below: the recorded layout has already drifted from the files once. What
follows was re-measured on `PSJCO204_REV指摘一覧_20260911.xlsx` on 2026-09-15; an older revision of
this section carried a different set of values taken from a `PSJCO308` copy (游ゴシック/Meiryo UI
9pt, merged `A1:M1`/`A2:M2`, header fill `12419407`, widths `5,8,13,8,24,30,62,40,11,10,10,11,26`,
severity fills `13551615`/`14083324`/`15922414`) and **none of those matched the PSJCO204 series**.
Measure, then copy.

One sheet named `REV指摘一覧`, no others. **The whole font is ＭＳ Ｐゴシック.** Row 1 is the title
(`<プログラムID>_<プログラム名>  REV指摘一覧`) in `A1`, **not merged**, 14pt bold, row height ~18.75.
Row 2 is the note line in `A2`, **not merged**, 9pt, carrying the target path, the checks actually
run, the REV date, and the sentence `ｾﾙ位置はA1形式(列記号+行番号)表記` — on a re-REV, also say which
prior findings the list does and does not repeat.

**Write `セル位置` in A1 notation (`AC571`, `B128`), not `[行,列]`.** The user asked for this
explicitly ("セル位置はA1とかの表記で出して欲しい") on the `PSJCO311` REV, and it is a standing
preference, not a one-off: reviewers navigate the workbook by typing the address into Excel's Name
Box, and `[571,29]` has to be converted by hand every time. The `[行,列]` form stays correct
*inside the chat report and in the sub-agents' own findings* — it is only the Excel 指摘一覧 that
takes A1. Convert at write time (column index → letter: repeatedly `((n-1) % 26)` → letter,
`n = (n-1) \ 26`), and when a finding spans several cells list them comma-separated
(`AC571, AC572, AC573`). Keep the sheet name in the separate `対象シート` column as before.

Row 3 is blank. Row 4 is the header, 10pt bold on fill
`14277593`, centred horizontally, top-aligned, wrapped:

`No. / 指摘ID / 分類 / 重要度 / 対象シート / セル位置 / 指摘内容 / 想定修正 / 対応要否 / 対応結果 / 対応者 / 対応日 / 備考`

Data starts at row 5, 10pt, wrapped, top-aligned; columns A (No.), B (指摘ID) and D (重要度) centred,
everything else left-aligned. **Fill only A-H and M — columns I-L (対応要否 / 対応結果 / 対応者 /
対応日) are left empty for the reviewer to fill in**; `備考` (M) is yours to write. Column widths
`5,8,15,8,30,34,66,44,11,39.83,10,11,26`. Thin borders (`LineStyle=1`, `Weight=2`) on every cell of
`A4:M<last>`, `AutoFilter` over the same range, and freeze panes below row 4.

`重要度` takes one of 高 / 中 / 低 / 要確認 and is coloured by **direct cell fill, not conditional
formatting** (the existing files have zero `FormatConditions`): 高 = `13421823` + bold, 中 =
`13434879`, 要確認 = `16772300`, 低 = left white. `分類` names the check the finding came from
(入出力定義 / 内部相互参照 / DBｶﾗﾑ / ID採番/記述ﾙｰﾙ / 更新条件表 / 帳票 / ﾌｧｲﾙ出力 / 誤字脱字 / 体裁), and `指摘ID` is
`<分類の連番>-<その中の連番>` (`2-13`), with column A a plain 1..N counter. Use the bare `体裁` the
existing files use — not `体裁(ﾌｫﾝﾄ)`/`体裁(結合)`, which this file previously specified and which no
copy of the list actually contains.

**`指摘ID` MUST be written as text, or Excel silently turns it into a date.** `1-1` becomes
`1月1日` and `2-13` becomes `2月13日` — every ID in the column, with no error and no visible warning
until someone opens the file. Set `NumberFormat = "@"` on the whole ID column **before** assigning
any value; setting it afterwards keeps the serial number and just reformats it. This is not
hypothetical — a run wrote all 49 rows this way and it was caught only by rendering the sheet to an
image afterwards.

**Do not rely on `AutoFit` for the data rows — set the row heights yourself.** Excel's `Rows.AutoFit()`
under-sizes wrapped Japanese text and silently clips the last line. Confirmed on the `PSJCO204`
20260915 list: after a range `AutoFit`, the final row's `指摘内容` lost its closing line, and it was
visible only in the PNG. Computing a height per row and applying it exactly still clipped that row
until a full line of slack was added. What works: for each wrapped column (C, E, F, G, H, M) measure
the text's display width — ASCII and **half-width katakana** (`U+FF61`-`U+FF9F`, which this project's
prose is full of) count 1, everything else 2 — take `lines = ceil(width / (columnWidth - 2))`, then
set `RowHeight = (max lines over those columns + 1) * 13.5`. The `+ 1` is the part that matters; the
rows come out slightly generous, which is the right trade for a list nobody should have to re-open
to read a truncated sentence.

Two PowerShell traps in the builder itself, both of which abort the run:

- **Declare the column-width array as `[double[]]`.** A bare `@(5,8,15,8,30,34,66,44,11,39.83,…)`
  is an `object[]` of mixed `int`/`double`, and assigning an element to `.ColumnWidth` throws
  `Specified cast is not valid.`
- **Never `Remove-Item` the output path to overwrite it.** `Remove-Item -LiteralPath $outPath -Force`
  is rejected by this environment's destructive-operation guard with the misleading
  `Remove-Item on system path '/' is blocked` — nothing runs at all. `SaveAs` with
  `$excel.DisplayAlerts = $false` overwrites an existing file without any prompt; just call it.

**The LAST `Chart.Export` in a loop comes back blank — always export one throwaway range after the
one you care about.** Confirmed twice in a single run on `PSJCO307`: a 4-band export loop wrote
bands 1-3 correctly and band 4 as a ~3KB blank PNG; re-running with three different bands wrote the
first two correctly and the third (the header band) as a ~1.5KB blank. The `Activate()` was present
every time, so this is not the missing-Activate failure below — it is the final iteration
specifically, apparently a clipboard/paste race as the loop unwinds. The fix is trivial: append a
small dummy range (a couple of rows) as the last entry of the band list and ignore its PNG. Check
every exported file's byte size before looking at it; anything under ~10KB for a populated range is
the blank.

**Verify by rendering, not by re-reading cell values.** The date bug is invisible to a `Value2`
dump. Export the finished range to PNG and actually look at it: `Range.CopyPicture(1,2)`, add a
`ChartObject` sized to `Range.Width`/`Range.Height`, `Activate()` it, `Chart.Paste()`, then
`Chart.Export(path,"PNG")` — the `Activate()` is required, without it the export writes a ~2.7KB
blank image that looks like a successful write. Split a long list into 2-3 ranges so the text stays
legible. PDF export works too but this machine has no `pdftoppm`/PyMuPDF, so a PDF cannot be viewed
back — PNG is the only route that closes the loop.

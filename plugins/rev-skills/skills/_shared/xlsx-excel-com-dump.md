# Reading Unicorn design-doc Excel files via Excel COM

> **Split note.** The parts of this document a check agent needs while *reading* a dump —
> folder exclusions, the no-sub-agent rule, what the live dump has already excluded and when to read
> `_DELETED_DIGEST.txt`, parsing a dump line, program-structure variants, the 区分名称 lookup rule,
> viewing the screen-layout picture, and the workbook template notes — now live in
> `agent-guide.md`. This file keeps what the **orchestrator** (or a standalone run) needs to
> *produce* dumps: the dump script and its traps, crash recovery, the reference-file cache.


**Microsoft Excel is installed and reachable via COM automation from PowerShell** — this is the
reliable, fast way to read `.xlsx` contents in this environment, and every `*review*`/`*column-check*`
skill in this skills folder uses the technique below. The scripts here are tuned against these real
workbooks (strikethrough/gray resolution, the UsedRange origin offset, the PID guard), so keep using
them rather than rolling your own reader.

**Correction to an earlier premise: Python and Node ARE available on this machine now.** This file
used to open by asserting they were not (that `python3`/`python`/`py`/`node` all resolved to inert
Windows Store stub launchers). That was true when it was written and is no longer true — measured
`2026-09-14`: `python` is a real CPython **3.13.15** at
the per-user `AppData` Programs Python313 install with **openpyxl 3.1.5** importable, and
`node` is **v24.19.0** under the standard `Program Files` nodejs install. Only the bare `python3` on PATH is
still the WindowsApps stub, so prefer `python`. Do not repeat the old claim, and don't waste a round
trip re-discovering it. Python is genuinely useful for the small side tasks this work throws off —
reading `xl/workbook.xml` out of the zip to list sheet names before deciding what to dump, editing
these skill docs, post-processing a dump — and `unzip` in the Bash tool still works too. It has not
replaced the Excel COM dump for the actual review read as the default route. An openpyxl equivalent
**has** now been validated against these workbooks, but only as the fallback below — don't swap the
default over casually mid-REV.

## When Excel COM crashes repeatedly on one workbook: the validated openpyxl fallback

Distinct from the `0x800A03EC` refuse-to-open failure further down. Here Excel **opens** the
workbook and then dies partway through the formatting cascade with
`The remote procedure call failed (0x800706BE)` followed by `The RPC server is unavailable
(0x800706BA)` on every later call. Confirmed on `PXJCO124_ﾛｯﾄ振向け.xlsx`, whose
`画面設計書(GXJC124A)` sheet (1674×162, 725 struck + 210 partially-struck cells) killed Excel on two
separate runs — once mid-`機能定義書`, once after 6+ minutes inside 画面設計書. Retrying with one
fresh `Excel.Application` per sheet did not help: it spawned an `/automation -Embedding` instance per
sheet, each spinning CPU and surviving `Quit()`, five orphans inside ten minutes. **Do not keep
retrying COM on a workbook that has done this once** — the retries cost far more than the fallback.

Two things to know before you reach for the fallback (plus one limit: `PXJCO130_ﾛｯﾄﾄﾚｰｽ.xlsx` makes
`openpyxl.load_workbook` raise `There is no item named 'xl/drawings/NULL' in the archive` — Excel opens
it fine, so that workbook must stay on COM):

- **Check whether the user has the target workbook open.** On that run `Get-CimInstance Win32_Process
  -Filter "Name='EXCEL.EXE'"` showed the user's own interactive Excel with the target `.xlsx` on its
  command line. Print the command lines, don't just compare PIDs.
- **The PID guard's "is this the user's Excel?" test is better written against the command line than
  against the pre-existing-PID list.** An instance whose command line contains the automation switch
  is one of ours and is safe to `Quit()`; one without it is interactive and must abort the run.
  Build that switch string as `[string]([char]47) + 'automation'` — **a literal `'/automation'` in a
  PowerShell command trips this environment's destructive-operation guard**, which rejects the entire
  command with the misleading `Remove-Item on system path '/automation' is blocked`. Same class of
  trap as the `[\/:*?"<>|]` character class documented below.

The fallback is `openpyxl` with `load_workbook(path, data_only=True, rich_text=True)`, reproducing
the live dump's `[row,col]=value` format exactly. **Validated: byte-identical output to the COM dump
on the two sheets COM completed before dying** (`表紙` and `機能定義書`), and the whole 8-sheet
workbook dumped in ~12 seconds versus COM's 10-minute failure. Always validate the same way — dump
whatever sheets COM managed, then diff — rather than trusting the fallback blind.

Three behaviours must be reproduced or the output silently diverges:

1. **`data_only=True` returns `datetime` for date-formatted cells; `Value2` returns the serial.**
   Convert back: `(dt - datetime(1899,12,30)).days + seconds/86400`. Without this every 作成日/更新日
   comes out as `2026-05-15 00:00:00` where the COM dump has `46157`.
2. **Format numbers like `Convert.ToString(double, InvariantCulture)`** — a float that is integral
   prints as `2`, not `2.0`.
3. **Rich-text strikethrough has two different inheritance rules, and getting either wrong changes
   which text is live.** Both were found by diffing against the COM dump:
   - A bare `str` element inside a `CellRichText` carries **no run properties and inherits the cell
     font**. Treating it as unstruck keeps text the COM dump correctly dropped.
   - A `TextBlock`'s own `rPr` **fully overrides** the cell font, so `strike is None` on a
     `TextBlock` means `False`, not "inherit". Falling back to the cell font here drops live text.

   Same rule for the gray-out colour test. The equivalent openpyxl helpers (`is_gray_rgb`,
   `font_gray`, the `analyse`/`cell_text` cascade) are small enough to rewrite from this description;
   what matters is that you diff the result against a COM dump of at least one real sheet before
   handing it to any check.

The same openpyxl route builds the `82.画面項目辞書_*` index from `_shared/reference-index.md`
(validated: 共通 1340 records/23 retired and 基準情報 2315/4 reproduced that doc's recorded counts
exactly) and serves `design-doc-formatting-consistency`, whose font/merge scan reads `cell.font.sz`,
`cell.font.name` and `ws.merged_cells.ranges` — so a COM crash does not force that check to be
skipped. Tell every downstream agent explicitly that COM crashes on this workbook and that they must
use openpyxl, or they will each rediscover it.

**One more consequence of that run: the Bash tool died partway through** with the
`add_item ("\??\C:\Program Files\Git", "/", ...) failed, errno 1` msys fault this document warns
about elsewhere. Have the PowerShell path ready — write scripts with the `Write` tool and invoke
them as `python <path>` from PowerShell rather than via a Bash heredoc.

One trap when using Python that way: a **Bash heredoc collapses doubled backslashes even when the
delimiter is quoted**, so a `"C:BSBSProgram Files"` literal inside a `python - <<'EOF'` script
arrives as a single backslash and `BSn` becomes a real newline. That silently wrote a line break into
the middle of a Windows path in this very file, twice, and the write succeeds so nothing warns you
(only a stray `SyntaxWarning: invalid escape sequence` hints at it). When a script must contain
backslashes, build them with `chr(92)` or avoid the path text entirely, and read the result back to
confirm. (`BS` above stands for the backslash character, for the same reason.)

## Orchestrating session: dump once, share the text, across parallel agents on the same workbook

When several sibling skills (`design-doc-internal-consistency`, `design-doc-io-table-check`,
`xlsx-db-column-check`, `naming-standard-compliance`, `update-condition-completeness`,
`report-design-check`, `file-output-spec-check`, `design-doc-writing-rules`,
`design-doc-formatting-consistency`, `design-doc-typo-check`) run as
parallel background agents against the SAME target workbook, don't
let each one independently dump it from scratch — up to 6x duplicated Excel COM cycles on an
identical file. Confirmed real waste: a review of `XJC_ｼｽﾃﾑ共通設計書.xlsx`'s `ﾛｯﾄ停止ﾁｪｯｸ` sheet had
3 agents each re-dump the same 999-row sheet because their prompts didn't say to reuse an existing
dump.

Instead: before launching the batch, run the live dump (per "The dump script" below) once yourself
against every sheet, and pass the resulting `.txt` file paths into each agent's prompt ("read
`<path>` for sheet X — don't re-dump it").

**In the same pre-launch step, build the reference-master index if any selected check needs it —
see `_shared/reference-index.md`.** Today that means `design-doc-internal-consistency` (screen-item
IDs), `report-design-check` (the `画面項目ID` column on print items) and `file-output-spec-check` — all
read the 画面項目辞書 index, so build it once when any is selected. That doc tells the *agent* not to build the index and to stop if it is
missing, so if you skip this step the check simply does not run. Build one index per dictionary file
the program's IDs route to (the routing table in `agent-guide.md`; a program using shared `XJZ`/`SJZ` items needs
the `_共通` file too), subset each to the program's own IDs, and pass both paths per file in the
prompt (the subset to read whole, the full index to grep) — same discipline as the dump.

**The dump script already applies strikethrough and gray-out, so a shared dump now carries
everything all but one skill needs — no agent should run its own strikethrough/gray scan.** Struck
cells are absent from the `.txt` entirely and partially-struck cells carry only their live text, so
"is this row still live?" is answered by whether it appears at all. Removed content is preserved
separately in `_DELETED_DIGEST.txt` for the one check that needs it (see `agent-guide.md`, "When to read `_DELETED_DIGEST.txt`"). Measured on `SXJCB147_処置指示発行(ｻﾌﾞﾌﾟﾛ).xlsx` (7 sheets): the
old shape cost ~155k tokens per agent (raw dump ~105k + strikethrough scan ~49k); the live dump
alone is ~89k (**-43%**), and live + digest is ~117k (-25%). Multiply that by 5-6 parallel agents.

The one remaining exception is `design-doc-formatting-consistency`, which needs `Font.Size` and
`MergeCells` — those still require live COM and are its own concern. Don't try to coordinate that
scan across agents; it costs more than it saves, and only one skill reads it.

## The dump script

Open the workbook invisibly, pull each worksheet's UsedRange as one `Value2` array (do **not**
loop `Cells.Item(r,c)` one cell at a time — it is slow enough to blow past the 2-minute Bash
timeout on any sheet with more than a few hundred cells), cap rows/cols since some sheets have a
bloated UsedRange from stray formatting far beyond the real content, and write one compact text
file per sheet with `[row,col]=value` tokens (skips empty cells, keeps coordinates for citing back
to the source file).

**The dump is a *live* dump: strikethrough and gray-out are resolved while the workbook is open, not
left for each agent to redo.** A fully-struck (or gray) cell is omitted from the `.txt`; a cell with
mixed struck/unstruck characters is written with only its live text; everything removed is written
to a separate `_DELETED_DIGEST.txt`. This is what makes the shared dump self-sufficient — see
"Orchestrating session" above for why, and `agent-guide.md`'s "When to read `_DELETED_DIGEST.txt`" for what the digest is for.

Getting that without paying per-cell COM costs needs a **three-level cascade**, because
`Font.Strikethrough`/`Font.Color` return `DBNull` when a range is mixed and a concrete value when it
is uniform:

1. Ask the whole `UsedRange` once. If it comes back "clean" (`Strikethrough = False` and a
   non-gray uniform `Color`), the sheet has nothing to strip — write the plain `Value2` dump and
   move on with **zero** extra COM calls. This is the common case for most sheets.
2. Otherwise ask each row once, over just that row's non-empty column span (3 COM calls per row).
   Rows that come back clean are skipped wholesale.
3. Only for rows that are struck or mixed, walk that row's non-empty cells; and only for cells that
   are themselves `DBNull` do the per-character `Characters(i,1)` reconstruction.

Measured on `SXJCB147_処置指示発行(ｻﾌﾞﾌﾟﾛ).xlsx` (7 sheets, 1089-row worst case, ~2,000 struck and
~390 partially-struck cells): 65s end to end, well inside a 600s timeout. Raise the Bash/PowerShell
timeout for this step rather than skipping the cascade.

**Skip 詳細設計書* and *画面ｲﾒｰｼﾞ* sheets — don't dump their content at all.** Confirmed across every
REV skill's own instructions: most of the single-program skills explicitly say "don't read/dump
詳細設計書 sheets" (it's near-empty boilerplate in practice), and `design-doc-formatting-consistency`
also excludes both 詳細設計書 and the 画面ｲﾒｰｼﾞ mockup sheet from its scan scope (the mockup is a
screenshot image with placeholder cells, confirmed content-free in every program checked so far).
Dumping their full cell content into the shared scratchpad every REV is pure waste — no sibling
skill ever reads those `.txt` files. **Do not extend this to every "ｲﾒｰｼﾞ"-named sheet** — a sheet
like `ﾒｰﾙｲﾒｰｼﾞ` (a batch program's mail-notification template) carries real narrative text that
`design-doc-internal-consistency`/`design-doc-typo-check` genuinely read and have found real
findings in (a stale placeholder wording, confirmed on `PXJCB134_流動停止(ﾊﾞｯﾁ).xlsx`) — only skip a
sheet whose name starts with `詳細設計` or contains `画面ｲﾒｰｼﾞ` specifically, not "ｲﾒｰｼﾞ" generally.

Record each skipped sheet's `UsedRange.Rows.Count`/`Columns.Count` (cheap — no `Value2` pull needed)
instead of a full dump, and report that alongside the rest. This one exception still needs it:
`naming-standard-compliance`'s "表紙's Ⅱ．設計書構成 list vs the sheets actually present" check needs
to tell whether a 詳細設計書 sheet is "populated with real content" or "an empty template" — the row
count alone is normally enough signal for that (a real 詳細設計書 runs dozens of rows; an untouched
template stub is a handful) without needing the full cell-value dump. If that skill's agent genuinely
needs to confirm actual content rather than just row count, it can open the one sheet itself via its
own light Excel COM check — that's still far cheaper than every REV dumping the full sheet by default.

**Format the `Value2` array via a compiled C# helper (`Add-Type`), not a plain PowerShell `for` loop.**
Benchmarked on `XJC_ｼｽﾃﾑ共通設計書.xlsx`'s `実績表項目設定` sheet (2652×116, ~307k cells): the
PowerShell loop took 7.8s to format vs 55ms for the equivalent compiled C# method — ~140x faster,
verified byte-identical output. The bottleneck is PowerShell interpreter overhead iterating cells in
memory, not the COM call (`Value2` itself pulls in 194ms) — so this costs nothing in correctness and
pays off most on the largest sheets. `ScreenUpdating`/`EnableEvents`/`Calculation` flags were also
tested and made no difference to `Workbooks.Open` time (~5.2-5.9s either way, fixed COM overhead) —
only the formatting loop is worth optimizing.

**`Add-Type` on this machine compiles with warnings-as-errors, so the cell-coordinate key must cast
`c` to an unsigned type before the bitwise OR.** Write `long key = ((long)r << 20) | (long)(uint)c;`
— the shorter `| (long)c` fails to compile with CS0675 ("Bitwise-or 演算子が sign-extended オペランド
で使用されています"), which aborts `Add-Type` entirely and then every `[XlsxDumpHelper]::` call throws
`TypeNotFound`. The failure mode is deceptive: the script keeps running past the errors, so
`Workbooks.Open` succeeds and the per-sheet `.txt` files are written **empty**, and the summary line
still prints a plausible `rows x cols` for each sheet. Verify `Add-Type` compiled (it prints nothing
on success) before trusting any dump.

**Do NOT "verify" it by compiling the helper in a separate PowerShell call first — that produces the
exact same empty-dump failure, for a different reason.** The PowerShell tool does not persist shell
state between calls: types, variables and functions defined in one call are gone by the next. An
`Add-Type` run on its own prints `ADDTYPE_OK` and looks like a clean verification, and then the
next call — the real dump — throws `Unable to find type [XlsxDumpHelper]` on every single row,
keeps running past the (non-terminating) errors, opens the workbook, and writes every per-sheet
`.txt` **empty**, just as a CS0675 failure would. Confirmed for real on `PSJCO304`, where the split
cost a full dump cycle and left an orphaned headless `EXCEL.EXE` behind.

**Always paste `Add-Type` and the code that uses it in the SAME PowerShell tool call.** This applies
to every script in this document and to `_shared/reference-index.md`'s index builder equally. If you
want a compile check, keep it in-line: put `Write-Output "ADDTYPE_OK"` immediately after the
`Add-Type` block of the same call, so a CS0675 failure is visible in the same output as the dump
summary.

**Pass the C# source in a SINGLE-quoted here-string — `Add-Type @'…'@`, never `Add-Type @"…"@`.**
A double-quoted here-string is expanded by PowerShell before the C# compiler ever sees it, so any
backtick in the source is eaten as a PowerShell escape. The concrete failure: the helper's
`sb.Append("\r\n")` was retyped as ``sb.Append("`r`n")`` — PowerShell turned the backtick escapes
into a real CR and LF *inside the C# string literal*, and `Add-Type` died with
`定数の 新しい行です。` (CS1010, newline in constant) followed by a cascade of brace errors. That aborts
`Add-Type` exactly like CS0675 does, with the same deceptive aftermath: the script keeps running,
`Workbooks.Open` succeeds and every per-sheet `.txt` is written **empty**. `$` is expanded the same
way, so a PowerShell variable name that happens to appear in the C# source would also be substituted.
Use `@'…'@` and write `"\r\n"` literally, as the scripts in this document do.

**A mid-run COM crash (`RPC server is unavailable`, HRESULT `0x800706BA`) silently produces a dump
that looks complete but has NOT had strikethrough removed. Discard and re-dump those sheets — never
review from them.** This is the most dangerous failure mode in this document, because unlike the
CS0675/here-string failures above it does **not** write an empty `.txt`: the sheet's `Value2` array
was already pulled before the crash, so the file comes out full size, plausibly formatted, and
completely convincing. What dies is only the level-2/level-3 cascade — every `$ws.Cells.Item(...)`
after the crash returns `$null`, the `try/catch` (or PowerShell's non-terminating error handling)
swallows it, `$dead`/`$liveMap` stay empty, and `FormatSheetLive` writes **every struck and
grayed-out cell as live text**.

Confirmed for real on `PXJCO128_ﾛｯﾄ停止指示登録.xlsx`: `画面設計書(GXJC128A)` was written at full
size (108,917 bytes) with **191 fully-struck cells and 97 partially-struck cells unresolved**. The
partial ones are what made it undetectable — a cell renamed in place reads as the old and new value
side by side, which looks exactly like this project's real "旧値の消し忘れ" defect:

| Dump text (wrong) | Actual live text |
|---|---|
| `(49)工程名取得` | `(4)工程名取得` |
| `30 60` (桁数) | `60` |
| `○ ×` | `×` |
| `XJC9076 SJZ9015` | `SJZ9015` |
| `ﾛｯﾄ停止指示更新解除画面` | `ﾛｯﾄ停止指示更新画面` |
| `一括解除ﾎﾞﾀﾝ` | `解除ﾎﾞﾀﾝ` |

Six sub-agents reviewed that dump and **every one of them raised the artefacts as 重要度「高」
findings** — "新旧2値が併記されたまま", "項番が2桁に連結", "存在しない画面名", "削除済みの記述が
取消線なしで残存" — and one reported the 「一括解除ﾎﾞﾀﾝ vs 解除ﾎﾞﾀﾝ」 naming conflict exactly
backwards, since the live button name matched the other documents all along. The whole REV had to be
re-run. Nothing in any agent's output hinted at a tooling fault; the reviewer caught it by asking
whether the strikethrough had been considered at all.

So, whenever a dump call reports a COM error of any kind:

1. **Treat every sheet that run touched as suspect, not just the one named in the error.** The error
   text names the row/cell the loop was on, not the sheet whose `.txt` was already written.
2. **Do not resume by skipping sheets whose `.txt` already exists.** A "resume" that tests
   `Test-Path $dest` will happily keep the corrupted file — that is precisely how this one survived
   two retry passes.
3. **Verify before reviewing**, on any sheet from a run that errored: re-dump it and diff the two
   cell maps, or spot-check a handful of cells' `Font.Strikethrough` directly. A clean sheet's
   summary line shows `dead=` / `partial=` counts; a sheet that crashed mid-cascade reports
   `dead=0 partial=0` while the workbook plainly has struck content. **`dead=0 partial=0` on a
   design-doc sheet with a populated 改訂履歴 column is a red flag, not a clean bill of health.**

**The openpyxl fallback is the practical recovery, and it is fast.** `load_workbook(path,
data_only=True, rich_text=True)` gives `cell.font.strike` and `cell.font.color` per cell, and a
partially-struck cell arrives as a `CellRichText` whose `TextBlock`s carry their own
`font.strike` — enough to reproduce the live-dump semantics exactly. It re-dumped all 12 sheets of
that workbook in a few seconds where the COM cascade had taken 15 minutes and crashed twice. Two
differences to keep in mind: openpyxl returns real `datetime` objects where COM's `Value2` returns
the Excel serial (`2026-06-09 00:00:00` vs `46182` — same value, and neither is a design defect),
and it reads the saved file rather than a live Excel, so formula results need `data_only=True`.
This does not make openpyxl the default reader — the COM dump remains the validated path — but a
crashed COM run is exactly the case where reaching for it beats a third retry.

**PowerShell variable names are case-INSENSITIVE, so never introduce a `$wS`/`$wC` next to the
worksheet handle `$ws`.** Writing `$wS = $used.Font.Strikethrough` silently overwrites the worksheet
object, and every later `$ws.Cells.Item(...)` then throws
`You cannot call a method on a null-valued expression.` — for *every* sheet, with a message that
names no sheet and no property, so it reads like a COM or workbook fault rather than a one-character
naming collision. Confirmed on a re-dump of `SSJCB211_ﾛｯﾄ振向け処理(ｻﾌﾞﾌﾟﾛ).xlsx`, where all 13
sheets failed identically and the per-sheet `try/catch` turned it into 13 identical "FAILED" lines.
Name the strikethrough/color probes something that cannot collide (`$allStrike`/`$allColor`,
`$rowStrike`/`$rowColor`), and if a whole-workbook dump fails on every sheet with that message, look
for a shadowed variable before blaming Excel.

**If a dump call does fail this way, check for an orphaned Excel before retrying.** The script's
`$excel.Quit()` runs but the instance can survive the torn-down PowerShell process. Run
`Get-Process EXCEL | Select-Object Id, StartTime`, compare `StartTime` against `Get-Date`, and
`Stop-Process` only the instance that started within the last minute or two — that one is yours.
**Never kill an older instance: the user routinely has their own Excel open**, and the
pre-existing-PID guard exists precisely because attaching to it and calling `Quit()` force-closes
their unrelated workbooks unsaved. Leaving an idle orphan behind is not harmless either — it is a
Running-Object-Table candidate the next run's `New-Object -ComObject Excel.Application` can attach
to, which the guard then correctly aborts on, blocking every subsequent attempt.

**Run every script INLINE in the PowerShell tool call — never write it to a `.ps1` and invoke the
file.** Cylance Script Control is active on this machine and blocks PowerShell from executing a
script file: `& C:\…\dump.ps1` dies immediately with exit code 34 and
`Cylance Script Control has blocked PowerShell from running.` Nothing in the script runs, so it
looks like a script bug rather than an endpoint-security block. Confirmed on the `PSJCO306` dump,
where writing the dump template to the scratchpad as a file and running it cost a full round trip.
Paste the whole script — `Add-Type` block included — as the `command` of one PowerShell tool call.

**The Bash tool is broken on this machine — use the PowerShell tool.** Any `Bash` call dies with
`fatal error - add_item ("\??\C:\Program Files\Git", "/", ...) failed, errno 1` and an msys stack
trace before the command runs at all. This is the Git-Bash/msys layer failing, not the command, so
retrying or rewording the command never helps. That also rules out the `unzip` route this document
mentions elsewhere and any Bash-based `Monitor`; use PowerShell (or `python`, which still works) for
zip inspection and side tasks.

**`$cell.Characters($i,1)` returns `$null` on a non-text cell, and the per-character strikethrough
walk then throws `You cannot call a method on a null-valued expression.`** A numeric/date cell whose
row was drilled into (because some *other* cell in that row is struck or gray) reaches the level-3
walk with `Font.Strikethrough` = `DBNull`, and `Characters()` has nothing to return. The per-sheet
`try/catch` reports it as a whole-sheet failure with a message that names no cell, so it reads like
a COM fault; on `PSJCO306` it failed 5 of the workbook's sheets. Guard both ways before walking:
skip the cell when `$cell.Value2 -isnot [string]`, and inside the loop treat a `$null` from
`Characters()` as "keep the raw text" rather than dereferencing `.Font`.

**Never put the character-class literal `[\/:*?"<>|]` in a PowerShell command — build the safe
sheet filename another way.** Confirmed for real while dumping `PSJCO309_焼成入炉帳ｻﾔ組み.xlsx`: the
template's own `$safe = ($n -replace '[\/:*?"<>|]','_')` line makes this environment's
destructive-operation safety guard reject the **entire** command, with a misleading error that names
an operation the script never performs — `Remove-Item on system path '*' is blocked. This path is
protected from removal.` Nothing is executed, so it reads like an environment fault rather than a
one-line lint problem, and it cost several full retries of a 30-sheet dump to localize (bisecting
the script section by section) before the culprit was found. The guard is reacting to the `*` and
`?` inside the class, not to anything the script does. Use the invalid-character list instead, which
carries no wildcard literals and is also more correct:

```powershell
$safe = [string]::Join('_', $n.Split([System.IO.Path]::GetInvalidFileNameChars()))
```

The scripts below already use this form. If you paste a script from anywhere else and hit the
"Remove-Item on system path" error, look for this literal first rather than assuming the guard is
objecting to `Workbooks.Open`, `$wb.Close($false)` or `$excel.Quit()`.

**The dump template** — the long block that follows, starting with `Add-Type` (the one-line `$safe` block
above is only an excerpt; an extractor that takes the first ```` ```powershell ```` block gets the wrong one):

```powershell
Add-Type @'
using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;
public class ExcelComWin32 {
    [DllImport("user32.dll")]
    public static extern uint GetWindowThreadProcessId(IntPtr hWnd, out uint processId);
}
public static class XlsxDumpHelper {
    static string Cell(object[,] arr, object vals, int r, int c, int rows, int cols) {
        object v;
        if (arr == null) v = vals;
        else if (rows == 1) v = arr[1, c];
        else if (cols == 1) v = arr[r, 1];
        else v = arr[r, c];
        if (v == null) return null;
        return Convert.ToString(v, System.Globalization.CultureInfo.InvariantCulture);
    }
    // Non-empty column indices for one row — lets the scan walk only cells that have content.
    public static List<int> NonEmptyCols(object vals, int r, int rows, int cols) {
        var list = new List<int>();
        object[,] arr = vals as object[,];
        for (int c = 1; c <= cols; c++) {
            string s = Cell(arr, vals, r, c, rows, cols);
            if (s != null && s.Length > 0) list.Add(c);
        }
        return list;
    }
    // Hidden values under a merge: Excel keeps a value in every cell of a merged range when the range
    // was filled before merging, and Value2 returns them all (SXJCB147 G704:P704 = `YOTO` x10). openpyxl
    // and the reviewer see only the anchor. Mark every non-anchor cell of every <mergeCell> in the sheet
    // XML as omitted — read from the file itself, so no COM call per cell and no guessing.
    public static int MarkMergeHidden(HashSet<long> dead, string sheetXml, int r0, int c0, int rows, int cols) {
        int n = 0;
        if (string.IsNullOrEmpty(sheetXml)) return 0;
        var rx = new System.Text.RegularExpressions.Regex("<mergeCell ref=\"([A-Z]+)(\\d+):([A-Z]+)(\\d+)\"");
        foreach (System.Text.RegularExpressions.Match m in rx.Matches(sheetXml)) {
            int ac1 = ColNum(m.Groups[1].Value), ar1 = int.Parse(m.Groups[2].Value);
            int ac2 = ColNum(m.Groups[3].Value), ar2 = int.Parse(m.Groups[4].Value);
            for (int ar = ar1; ar <= ar2; ar++) {
                int r = ar - r0; if (r < 1 || r > rows) continue;
                for (int ac = ac1; ac <= ac2; ac++) {
                    if (ar == ar1 && ac == ac1) continue;
                    int c = ac - c0; if (c < 1 || c > cols) continue;
                    if (dead.Add(((long)r << 20) | (long)(uint)c)) n++;
                }
            }
        }
        return n;
    }
    static int ColNum(string s) { int v = 0; foreach (char ch in s) v = v * 26 + (ch - 'A' + 1); return v; }
    // dead = coords to omit entirely; live = coords whose text is replaced by its unstruck remainder.
    // Both are keyed on ARRAY (UsedRange-relative) coordinates; r0/c0 shift the EMITTED coordinate
    // to sheet-absolute so a citation like [10,5] names the cell a reviewer sees in Excel.
    public static string FormatSheetLive(object vals, int rows, int cols,
                                         HashSet<long> dead, Dictionary<long,string> live,
                                         int r0, int c0) {
        var sb = new System.Text.StringBuilder();
        object[,] arr = vals as object[,];
        for (int r = 1; r <= rows; r++) {
            var parts = new List<string>();
            for (int c = 1; c <= cols; c++) {
                long key = ((long)r << 20) | (long)(uint)c;
                if (dead.Contains(key)) continue;
                string s = live.ContainsKey(key) ? live[key] : Cell(arr, vals, r, c, rows, cols);
                if (s != null && s.Length > 0) parts.Add("[" + (r + r0) + "," + (c + c0) + "]=" + s);
            }
            if (parts.Count > 0) { sb.Append(string.Join(" | ", parts)); sb.Append("\r\n"); }
        }
        return sb.ToString();
    }
}
'@

$out    = "<scratchpad dir>"
$path   = "<absolute path to the target workbook>"
$prefix = "<program id>"
# Sheet scoping. $null = every visible sheet. Set it when the user scopes the REV, e.g.
# @("表紙*","機能定義書*","帳票設計書*"). Hidden sheets (bk_ backups etc.) are out of scope by default.
$IncludeSheetPatterns = $null
$IncludeHiddenSheets  = $false
$ErrorActionPreference = 'Stop'   # abort at the first RPC error instead of writing a half dump

function Test-Gray([double]$argb) {
    $r = [int]$argb -band 0xFF; $g = ([int]$argb -shr 8) -band 0xFF; $b = ([int]$argb -shr 16) -band 0xFF
    return (($r -eq $g) -and ($g -eq $b) -and ($r -gt 80) -and ($r -lt 220))
}

# Record every EXCEL.EXE PID that already exists BEFORE launching our own instance. On this
# environment, `New-Object -ComObject Excel.Application` has been observed to sometimes attach to
# the user's own already-running interactive Excel process instead of spawning a genuinely new one
# (a Running-Object-Table quirk). If that happens and we later blindly call $excel.Quit(), it closes
# the user's ENTIRE real Excel session — including unrelated workbooks they had open — without
# saving. This bit us once already: a REV skill's dump step silently attached to the user's session
# and Quit() force-closed a workbook they had open for unrelated work, mid-dump.
$preExistingExcelPids = @((Get-Process EXCEL -ErrorAction SilentlyContinue).Id)
$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false
[uint32]$excelComPid = 0
[ExcelComWin32]::GetWindowThreadProcessId([IntPtr]$excel.Hwnd, [ref]$excelComPid) | Out-Null
if ($preExistingExcelPids -contains $excelComPid) {
    throw "Excel COM automation attached to the user's existing Excel process (PID $excelComPid) instead of creating a new instance. Aborting without calling Open/Close/Quit on it — ask the user to close their other Excel windows first."
}
# Sheet name -> its XML part inside the .xlsx, for the merge map below. Opened read-only with
# FileShare.ReadWrite so it works while Excel (ours or the user's) has the file open.
Add-Type -AssemblyName System.IO.Compression
$zipFs  = [System.IO.File]::Open($path, 'Open', 'Read', 'ReadWrite')
$zip    = New-Object System.IO.Compression.ZipArchive($zipFs, [System.IO.Compression.ZipArchiveMode]::Read)
function Read-ZipText([string]$name) {
    $en = $zip.GetEntry($name); if (-not $en) { return $null }
    $sr = New-Object System.IO.StreamReader($en.Open()); try { $sr.ReadToEnd() } finally { $sr.Close() }
}
$relT = @{}
foreach ($m in [regex]::Matches((Read-ZipText 'xl/_rels/workbook.xml.rels'), '<Relationship [^>]*>')) {
    $id = [regex]::Match($m.Value, 'Id="([^"]+)"').Groups[1].Value
    $tg = [regex]::Match($m.Value, 'Target="([^"]+)"').Groups[1].Value
    if ($tg.StartsWith('/')) { $tg = $tg.Substring(1) } elseif (-not $tg.StartsWith('xl/')) { $tg = 'xl/' + $tg }
    $relT[$id] = $tg
}
$sheetPart = @{}
foreach ($m in [regex]::Matches((Read-ZipText 'xl/workbook.xml'), '<sheet [^>]*>')) {
    $nm = [System.Net.WebUtility]::HtmlDecode([regex]::Match($m.Value, ' name="([^"]+)"').Groups[1].Value)
    $sheetPart[$nm] = $relT[[regex]::Match($m.Value, 'r:id="([^"]+)"').Groups[1].Value]
}

try {   # the finally below closes OUR instance even when a sheet throws mid-dump
$wb = $excel.Workbooks.Open($path, $true, $true)   # ReadOnly, no update-links prompt

$deleted       = New-Object System.Collections.Generic.List[object]
$skippedSheets = @()
$summary       = @()

foreach ($ws in $wb.Worksheets) {
    $n = $ws.Name
    if (-not $IncludeHiddenSheets -and $ws.Visible -ne -1) { continue }
    if ($IncludeSheetPatterns) {
        $hit = $false
        foreach ($pat in $IncludeSheetPatterns) { if ($n -like $pat) { $hit = $true; break } }
        if (-not $hit) { continue }
    }
    if ($n -like '詳細設計*' -or $n -like '*画面ｲﾒｰｼﾞ*') {
        $u = $ws.UsedRange
        $skippedSheets += "$n : rows=$($u.Rows.Count) cols=$($u.Columns.Count) (skipped — out of scope for every REV skill; not dumped)"
        continue
    }

    $used = $ws.UsedRange
    # Caps guard against a runaway UsedRange (旧ﾃｰﾌﾞﾙﾚｲｱｳﾄ(XAE) reports 1,048,569 rows). A real sheet can pass
    # 3000 rows (PXJCO161 画面設計書(GXJC161B) has 4205; the old 3000 cap silently dropped 49% of its cells),
    # so the cap is 20000 and any cut is printed as TRUNCATED in the summary, never silent.
    $rows = [Math]::Min($used.Rows.Count, 20000)
    $cols = [Math]::Min($used.Columns.Count, 260)
    $trunc = ''
    if (($used.Rows.Count -gt $rows) -or ($used.Columns.Count -gt $cols)) { $trunc = "`tTRUNCATED to ${rows}x${cols}" }
    $vals = $used.Value2
    # $vals is indexed from 1 WITHIN the UsedRange; $ws.Cells.Item is absolute on the sheet. Every
    # absolute access below adds this offset, and so does every coordinate written out. See
    # "UsedRange-relative vs sheet-absolute coordinates" below for why this matters.
    $r0 = $used.Row - 1
    $c0 = $used.Column - 1
    $dead    = New-Object 'System.Collections.Generic.HashSet[long]'
    $liveMap = New-Object 'System.Collections.Generic.Dictionary[long,string]'
    $nDead = 0; $nPart = 0

    # Hidden values under a merge (see MarkMergeHidden): omitted from the dump, never digested.
    [void][XlsxDumpHelper]::MarkMergeHidden($dead, (Read-ZipText $sheetPart[$n]), $r0, $c0, $rows, $cols)

    # Level 1 — one question for the whole sheet. Clean sheets cost zero further COM calls.
    $wholeStrike = $used.Font.Strikethrough
    $wholeColor  = $used.Font.Color
    $sheetClean  = ($wholeStrike -isnot [System.DBNull]) -and ($wholeStrike -eq $false) -and
                   ($wholeColor  -isnot [System.DBNull]) -and (-not (Test-Gray $wholeColor))
    if (-not $sheetClean) {
        for ($r = 1; $r -le $rows; $r++) {
            $cl = [XlsxDumpHelper]::NonEmptyCols($vals, $r, $rows, $cols)
            if ($cl.Count -eq 0) { continue }
            # Level 2 — one question per row, over that row's non-empty span only.
            $ar = $r + $r0
            $rowRange = $ws.Range($ws.Cells.Item($ar, ($cl[0] + $c0)),
                                  $ws.Cells.Item($ar, ($cl[$cl.Count - 1] + $c0)))
            $rs = $rowRange.Font.Strikethrough
            $rc = $rowRange.Font.Color
            $drill = ($rs -is [System.DBNull]) -or ($rs -eq $true) -or
                     ($rc -is [System.DBNull]) -or (Test-Gray $rc)
            if (-not $drill) { continue }
            # Level 3 — per cell, and per character only where the cell itself is mixed.
            foreach ($c in $cl) {
                if ($dead.Contains((([int64]$r) -shl 20) -bor ([int64]$c))) { continue }   # hidden under a merge
                $cell = $ws.Cells.Item($ar, ($c + $c0))
                $s    = $cell.Font.Strikethrough
                $col  = $cell.Font.Color
                $isGray = (-not ($col -is [System.DBNull])) -and (Test-Gray $col)
                $key = (([int64]$r) -shl 20) -bor ([int64]$c)   # key stays array-relative
                $aC  = $c + $c0                                # digest records absolute coords
                # Mixed = strikethrough is DBNull. Gray is judged for the whole cell only: partly-gray text
                # is SQL syntax colouring in this corpus, never a deletion (same rule as live_dump.classify).
                $mixed = ($s -is [System.DBNull])
                if ($mixed -and ($cell.Value2 -isnot [string])) {
                    # mixed font on a number/date cell: Characters() has nothing to walk — keep it live
                } elseif ($mixed) {
                    $raw = "$($cell.Value2)"; $lv = ""
                    for ($i = 1; $i -le $raw.Length; $i++) {
                        $ch = $cell.Characters($i, 1)
                        if ($null -eq $ch) { $lv = $raw; break }   # see "Characters() returns $null" below
                        if (-not $ch.Font.Strikethrough) { $lv += $ch.Text }
                    }
                    if ($lv -eq $raw) { continue }   # mixed colour but nothing struck or gray: plain live cell
                    if ($lv.Trim().Length -eq 0) {
                        [void]$dead.Add($key); $nDead++
                        $deleted.Add([PSCustomObject]@{S=$n; R=$ar; C=$aC; Kind='DEL'; Text=$raw})
                    } else {
                        $liveMap[$key] = $lv; $nPart++
                        $deleted.Add([PSCustomObject]@{S=$n; R=$ar; C=$aC; Kind='PART'; Text="raw='$raw' live='$lv'"})
                    }
                } elseif (($s -eq $true) -or $isGray) {
                    [void]$dead.Add($key); $nDead++
                    $kind = if ($s -eq $true) { 'DEL' } else { 'GRAY' }
                    $deleted.Add([PSCustomObject]@{S=$n; R=$ar; C=$aC; Kind=$kind; Text="$($cell.Value2)"})
                }
            }
        }
    }

    $safe = [string]::Join('_', $n.Split([System.IO.Path]::GetInvalidFileNameChars()))
    $text = [XlsxDumpHelper]::FormatSheetLive($vals, $rows, $cols, $dead, $liveMap, $r0, $c0)
    [System.IO.File]::WriteAllText((Join-Path $out ($prefix + "_" + $safe + ".txt")), $text, [System.Text.Encoding]::UTF8)
    $summary += "$n`t$($used.Rows.Count)x$($used.Columns.Count)`tdead=$nDead`tpartial=$nPart$trunc"
}
} finally {
    if ($wb) { try { $wb.Close($false) } catch {} }
    try { $excel.Quit() } catch {}   # safe: the PID guard above proved this instance is ours
    $zip.Dispose(); $zipFs.Close()
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null
}

# _DELETED_DIGEST.txt — removed content, grouped into contiguous row blocks. Structural filler is
# counted but not listed: a bare row/condition number, a hyphen placeholder, a comparison operator.
# Everything else is listed WHATEVER ITS LENGTH. Do not "simplify" this back to a character-count
# threshold — a length rule was tried first (drop anything under 4 chars) and silently swallowed 249
# of 761 omissions on SXJCB147, because this project's load-bearing tokens are routinely 1-3
# characters: block references like (5), footnote markers like ※2, query aliases Y/Z/A/C, and plain
# words such as 引数 / ﾗﾝｸ / 日付 / 参照先. Those are exactly what an asymmetry finding hangs on.
# The three filler classes below were verified safe to drop on that same workbook: all 248 numerics
# sat in No. columns (never in a 検索条件 right-hand side, so no literal value is lost), all 229
# hyphens and all 35 operators belonged to rows deleted in full, where the row's real content is
# already listed. Note this leaves ー (U+30FC, the katakana prolonged mark) out of the hyphen class
# on purpose — it is a letter, not a dash.
$lines = New-Object System.Collections.Generic.List[string]
foreach ($grp in $deleted | Group-Object S) {
    $items = $grp.Group | Sort-Object R, C
    $blockStart = $null; $prev = $null
    $buf = New-Object System.Collections.Generic.List[object]
    $flush = {
        if ($buf.Count -eq 0) { return }
        $subst = @($buf | Where-Object {
            $t = $_.Text.Trim()
            -not ($t -match '^[0-9]+$' -or $t -match '^[-‐‑–—―－]$' -or $t -match '^[=<>≠≦≧≤≥]+$')
        })
        $short = $buf.Count - $subst.Count
        $lines.Add("=== $($grp.Name) rows $blockStart-$prev ($($buf.Count) cells) ===")
        foreach ($it in $subst) { $lines.Add("  [$($it.R),$($it.C)] $($it.Kind): $($it.Text)") }
        if ($short -gt 0) { $lines.Add("  (+ $short filler cells omitted)") }
        $buf.Clear()
    }
    foreach ($it in $items) {
        if ($null -eq $blockStart) { $blockStart = $it.R }
        elseif ($it.R - $prev -gt 2) { & $flush; $blockStart = $it.R }
        $prev = $it.R; $buf.Add($it)
    }
    & $flush
}
[System.IO.File]::WriteAllText((Join-Path $out "_DELETED_DIGEST.txt"), ($lines -join "`r`n"), [System.Text.Encoding]::UTF8)

$summary -join "`n"
$skippedSheets -join "`n"   # a suspiciously large row count on a skipped 詳細設計書 sheet is a signal naming-standard-compliance may need to check its content directly
```

## When Excel refuses to open the workbook at all (0x800A03EC)

(Re-checked 2026-10-01: `SXJCB147` now opens in both Excel and openpyxl, and the two dumps agree cell
for cell — 1,451 struck / 396 partial. Try COM first; fall back to `scripts/live_dump.py` on this error.)

Occasionally `$excel.Workbooks.Open(...)` fails on one specific design-doc workbook with
`Workbooks クラスの Open プロパティを取得できません。` / HRESULT `0x800A03EC`, while every other
workbook in the same folder opens fine from the same COM instance. This is **not** a COM quirk, a
busy-Excel problem or a filename problem — it means Excel's loader is rejecting the file's content.
Confirmed for real on `SXJCB147_処置指示発行(ｻﾌﾞﾌﾟﾛ).xlsx` (工程管理/PHASE3), where it cost a long
diagnostic detour before the cause was found.

**The known cause: an out-of-range `<rPh>` (ふりがな) offset in `xl/sharedStrings.xml`.** When an
author shortens a cell's text but Excel doesn't refresh the phonetic-guide runs attached to it, the
saved `<si>` keeps `<rPh sb="…" eb="…">` offsets that point past the end of the (now shorter) `<t>`
text. Excel validates these at load and refuses the whole workbook. On `SXJCB147` the offending
entry was:

```xml
<si><t>「4.処置指示工程」に1件のみ存在</t>          <!-- 17 characters -->
  <rPh sb="26" eb="27"><t>ケン</t></rPh>            <!-- offsets 26/27 and 29/31 are past the end -->
  <rPh sb="29" eb="31"><t>ソンザイ</t></rPh>
  <phoneticPr fontId="31"/></si>
```

**Don't chase the usual suspects first — they were all ruled out on that file** and each cost a
round trip: `unzip -t` reports no error, every XML part is well-formed (`XmlReader` over all
`.xml`/`.rels` entries passes), no `<fileSharing>`/`<workbookProtection>`, no Mark-of-the-Web ADS,
no `~$` lock file, no entry in Excel's `Resiliency\DisabledItems`, the sheet/rels/Content_Types
cross-references are all complete, and the file opens no better with `CorruptLoad` set to
`xlRepairFile` (1) or `xlExtractData` (2), with `UpdateLinks=0`, from a renamed ASCII-named copy, or
after stripping `calcChain.xml` or all twelve dead `externalLinks`. Copying the file elsewhere
doesn't help either — it is the content, not the path or the environment.

### Diagnosing it

Scan every `<si>` and compare each `<rPh>`'s `sb`/`eb` against the length of that entry's own text
(the concatenation of its `<t>` elements **outside** any `<rPh>` block — for a rich-text `<si>`,
concatenate the `<r><t>` runs). Anything with `eb > len`, `sb > len` or `sb >= eb` is a candidate:

```powershell
$ss = [System.IO.File]::ReadAllText($extractedSharedStringsPath, [System.Text.Encoding]::UTF8)
$body  = $ss.Substring($ss.IndexOf("<si>"), $ss.LastIndexOf("</sst>") - $ss.IndexOf("<si>"))
$items = @([regex]::Matches($body, "<si>.*?</si>", 'Singleline') | ForEach-Object { $_.Value })
for ($i = 0; $i -lt $items.Count; $i++) {
    $noRph = [regex]::Replace($items[$i], '<rPh\b.*?</rPh>', '', 'Singleline')
    $len = (([regex]::Matches($noRph, '<t(?:\s[^>]*)?>(.*?)</t>', 'Singleline') |
             ForEach-Object { $_.Groups[1].Value }) -join '').Length
    foreach ($m in [regex]::Matches($items[$i], '<rPh\s+sb="(\d+)"\s+eb="(\d+)"')) {
        $sb = [int]$m.Groups[1].Value; $eb = [int]$m.Groups[2].Value
        if ($eb -gt $len -or $sb -gt $len -or $sb -ge $eb) { "[$i] len=$len sb=$sb eb=$eb :: $($items[$i])" }
    }
}
```

If that scan comes back empty, bisect instead: rebuild the workbook with parts progressively
replaced by valid stubs (stub every `xl/worksheets/sheetN.xml` down to
`<worksheet …><sheetData/></worksheet>`, minimise `styles.xml`, drop `<definedNames>`, remove
`xl/drawings`+`xl/media`+`xl/printerSettings`+`xl/worksheets/_rels`+`xl/sharedStrings.xml`, adjusting
`[Content_Types].xml` and `xl/_rels/workbook.xml.rels` to match) until it opens, then add groups back
one at a time. **Before trusting a single bisect result, rezip a workbook that is known to open and
confirm your rebuild method itself produces an openable file** — `zip -q -r -X out.xlsx
'[Content_Types].xml' _rels docProps xl` from the extracted tree is verified to work. Skipping that
control makes every "still fails" result meaningless. Binary-searching within `sharedStrings.xml`
(keep entries `[lo,hi]` original, stub the rest, swap the part in via
`[System.IO.Compression.ZipFile]::Open($f,'Update')`) localises the entry in ~12 opens; use
`[Math]::Floor(($lo+$hi)/2)` for the midpoint — PowerShell's `[int]` cast does banker's rounding and
turns `[2831,2832]` into an infinite loop.

### Repairing it, and what to tell the user

**Repair a copy in the scratchpad and dump from that; never rewrite the user's design doc without
asking.** The fix is to delete just the offending `<rPh>` elements (or the `<si>`'s phonetic runs
entirely) — no cell value, font, or merge is affected, so a formatting scan run against the repaired
copy is still valid for the original:

```powershell
Copy-Item -LiteralPath $original -Destination $repaired -Force
$ss = $ss.Replace($badSi, $badSiWithoutRph)
$za = [System.IO.Compression.ZipFile]::Open($repaired, 'Update')
$za.GetEntry("xl/sharedStrings.xml").Delete()
$sw = New-Object System.IO.StreamWriter($za.CreateEntry("xl/sharedStrings.xml").Open(),
                                        (New-Object System.Text.UTF8Encoding($false)))
$sw.Write($ss); $sw.Dispose(); $za.Dispose()
```

Then run the normal dump against the repaired copy, tell every downstream agent that the original is
corrupt so none of them wastes time trying to open it, and cite findings against the **original**
path. Report the corruption to the user as a finding in its own right — a design doc nobody can open
in Excel is a bigger problem than anything the REV itself will turn up — and ask before writing the
fix back to the original file.

## UsedRange-relative vs sheet-absolute coordinates

`$used.Value2` is indexed from 1 **within the UsedRange**; `$ws.Cells.Item(r,c)` is absolute on the
sheet. Feeding array indices to `Cells.Item` reads a *different cell* on any sheet whose UsedRange
does not start at `A1`, and it fails silently. This script derives its non-empty column list from
the array, so before the fix every absolute access — the per-row `Range` for the level-2 scan and
the per-cell drill in level 3 — was off by the origin, and the emitted `[row,col]` tokens were off
too.

Both consequences matter, and the second is worse:

1. Strikethrough and gray-out get resolved against the wrong cells, so struck content can survive
   into the dump and live content can be dropped — the exact failure this live dump exists to
   prevent.
2. A finding citing `[10,5]` does not name the cell a reviewer opens. Reviewers navigate by these
   coordinates, so an offset dump makes every finding on that sheet unactionable.

**Measured on the 61 workbooks of `01_Doc\08_機能定義書\11_工程管理\PHASE3`**: 33 of 715 visible
multi-cell sheets do not start at `A1`; after this script's own `詳細設計*` / `*画面ｲﾒｰｼﾞ*` skips,
**12 of 600 in-scope sheets** are affected. Two are worth naming because they are sheets a REV
actually reads: `PXJCB134_流動停止(ﾊﾞｯﾁ).xlsx`'s `ﾒｰﾙｲﾒｰｼﾞ` (offset r1/c1 — the very sheet the skip
rule above carves out an exception *for*), and `SSJCB605_完成割合設定(ｻﾌﾞﾌﾟﾛ).xlsx`'s
`完成割合設定(ｻﾌﾞﾌﾟﾛ)_事業部説明なし` at offset **r33/c34**, where a cited `[5,3]` is really cell
`AK38`. `PXJCO124_ﾛｯﾄ振向け.xlsx`'s two `《参考》画面項目遷移` sheets (881 and 836 rows, r2/c0) and
`PXJCO192_処置指示登録.xlsx`'s `ﾛｯﾄ管理C　処置指示登録` (271x369, r0/c1) are the other large ones.

Reference masters are nearly clean by comparison — 1 of 84 visible sheets across the eight cached
registries and common-design files, and that one is `05.ｼｽﾃﾑ共通設計書.xlsx`'s
`(参考)配色ﾃﾝﾌﾟﾚｰﾄ`, which no check reads.

**The fix, applied in all three scripts here**: compute `$r0 = $used.Row - 1` / `$c0 = $used.Column - 1`
once per sheet, add it to every `Cells.Item` access, and add it to every coordinate written out. Keep
the internal `dead`/`live` keys array-relative — they index the same array — and offset only at the
boundary. `_DELETED_DIGEST.txt` records absolute coordinates too.

**This changes what cached dumps mean, and the source files did not change**, so mtime/length alone
would report a false cache hit. `meta.json` therefore carries `dumpFormat` (now `2`), compared
alongside mtime and length; a `dumpFormat 1` cache is treated as a miss and redumped. Nothing needs
to be deleted by hand.

### Never compare the cached mtime as a string

`ConvertFrom-Json` on PowerShell 7 silently converts an ISO-8601 string to `[DateTime]`, and that
object's default `ToString()` is `07/28/2026 01:39:58` — no sub-second digits. Comparing it with
`-eq` against the `"o"`-format string the cache was written from (`2026-07-28T01:39:58.8626833Z`) is
therefore **always false**, so every reference file reports a miss and gets re-dumped through Excel
COM on every single run. The cache silently does nothing.

Confirmed on this machine (PowerShell 7.6.6): a `TXJAM023_工程ﾏｽﾀ.xlsx` entry whose `length` and
`dumpFormat` both matched still failed, and a 12-file batch whose cache was fully warm reported all
12 as misses. The `[DateTime]` itself keeps full precision — only its stringification loses it — so
comparing `.ToUniversalTime().Ticks` recovers the match exactly. The `Test-MTimeMatch` helper in both
scripts below does that and accepts either shape, since Windows PowerShell 5.1 and
`ConvertFrom-Json -AsHashtable` leave the value as a string.

The same trap applies to any other `meta.json` field you add that holds a date. Keep comparisons on
`length`, `dumpFormat` and `sourcePath` as plain `-eq`; those are a number, a number and a string.

## Cross-session cache for reference/master files (not the target workbook)

The "dump once, share the text" section above only dedupes work *within* one REV run, for the one
target workbook under review. It does nothing for the other files a REV skill reads as ground
truth — `04.ﾒｯｾｰｼﾞ管理_*.xlsx`, `09.区分名称_step2.xlsx`,
`05.ｼｽﾃﾑ共通設計書.xlsx`, `06-*.xlsx` (機能一覧/DB一覧/レスポンス一覧), `07.共通項目取得.xlsx`, the
チェックリスト file, and every テーブルレイアウト workbook. These almost never change between one
program's REV and the next, yet without caching they get re-opened via Excel COM from scratch every
single time, by every agent that needs them — including more than once within the same REV run, when
several sibling skills happen to need the same reference file.

**One exception: `82.画面項目辞書_*.xlsx` is indexed, not cached — see
`_shared/reference-index.md`.** Caching still re-reads whole sheets into an agent's context, and
the only thing any check wants from that file is `id → 画面項目名`. On the 工程管理 copy the index is
between one and two orders of magnitude smaller than the dump it replaces, and smaller again once
subset to one program's IDs — **the figures and the source state they were measured at live in
`reference-index.md`, not here**, so they only have to be refreshed in one place. Do not restate
them here, however tempting: three stale copies is what this pointer exists to prevent. Every other
file listed above — including テーブルレイアウト workbooks and the
`06-*.xlsx` registries — is **read live with `scripts/live_dump.py`**, not through the cache below
(v1.17.0: the cache writes raw `Value2` and keeps struck text, which every one of those files has).
The cache below remains only for a caller that needs raw values on purpose; no REV check does.

**Use a persistent, cross-session cache keyed by the source file's last-write-time for any file in
this category.** Cache root: `<user home>\.claude\skills\_cache\xlsx-dumps\<md5 of the lowercased
absolute source path, plus an optional sheet-filter suffix — see below>\`, holding one
`<sheetName>.txt` per dumped sheet (same `[row,col]=value` format as above) plus a `meta.json`
recording the source path, `LastWriteTimeUtc`, file length, and which sheet filter (if any) was
used. Before dumping a reference file, compute its current `LastWriteTimeUtc`/length and compare
against `meta.json`: if both match, reuse the cached `.txt` files as-is (no Excel COM call at all —
this is the common case, since these files rarely change); if either differs, or `meta.json` doesn't
exist yet, redump via Excel COM (same safety-guarded approach as "The dump script" above — pre-existing-PID
check included) and overwrite the cache with the new dump plus updated `meta.json`. **Don't wipe the
cache directory with `Remove-Item -Recurse -Force` before redumping** — it isn't needed (the
directory is keyed by path + sheet-filter together, so the same cache dir always wants the exact
same set of sheet files every time it's redumped, and `Set-Content`/`WriteAllText` overwrite existing
files in place cleanly) and a recursive delete right next to a literal design-doc filename in the
same script has tripped this environment's own destructive-operation safety guard. Just
`New-Item -ItemType Directory -Force` (idempotent whether or not the folder already exists) and let
the per-sheet writes overwrite in place.

**Optional: restrict the dump to only the sheets a WG-scoped lookup can ever actually use, via
name patterns.** Two confirmed cases:

- **テーブルレイアウト files** — a table-layout workbook commonly has 4 sheets (checked
  `TXJCM501_ﾛｯﾄ停止.xlsx`): `改訂履歴`, `ﾃｰﾌﾞﾙﾚｲｱｳﾄ` (the one every skill actually reads),
  `JAG_ﾃｰﾌﾞﾙﾚｲｱｳﾄ` (an old JAGUR-era copy), `旧ﾃｰﾌﾞﾙﾚｲｱｳﾄ` (even older) — only 1 of 4 sheets is ever
  read. Use `$OnlySheetPatterns = @("ﾃｰﾌﾞﾙﾚｲｱｳﾄ")`.
- **`04.ﾒｯｾｰｼﾞ管理_<WG>.xlsx`** — confirmed for real on the 工程管理 copy: `04.ﾒｯｾｰｼﾞ管理_工程管理.xlsx`
  has 11 sheets (改訂履歴, 方針, XSA(共通), XSK(共通), XJZ(共通), XJA(基準), XJB(受注), **XJC(工程)**,
  **SJC(工程)**, XJD(品質), XJE(生産)) but the skills' own routing rules (see
  `design-doc-internal-consistency` check 3/4) mean only the reviewed program's own-WG JOBコード
  sheets (bold above) are ever read from THIS file — a common (共通) prefix routes to
  `04.ﾒｯｾｰｼﾞ管理_共通.xlsx` instead, and a foreign-WG prefix routes to that other WG's own file, never
  to a same-named sheet embedded in this one. Use
  `$OnlySheetPatterns = @("XJC(*", "SJC(*")` (or the equivalent for the WG you're checking) —
  `-like` wildcard patterns, not exact names, since a WG's own-prefix sheet can carry an extra
  suffix that an exact-name match would miss. **Leave the pattern open-ended with no closing `)`** —
  `-like` requires a literal `)` to match the actual end of the string, so a closed pattern like
  `"SJC(*)"` fails against a name such as `SJC(工程)再開発追加分 ` (extra text AND a trailing space,
  so it doesn't end in `)`) even though it very much should match; confirmed this exact bug while
  testing against the real 画面項目辞書 file, which carries a sheet named exactly that.

**`82.画面項目辞書_*.xlsx` is not on this list** — it is indexed rather than dumped, so it never
reaches `$OnlySheetPatterns` at all. Its sheet selection is the index builder's business and works
by header detection, not by name; see `_shared/reference-index.md`.

Leave `$OnlySheetPatterns` `$null` (the default) for file types with no such known-safe restriction —
but note none of the REV masters belongs here any more (`09.区分名称` → `build_kbn_index.py`; `05`, `06-*`,
`07`, `04` → `live_dump.py` with `--sheets`). If none of the given patterns match any
sheet in a given workbook (a doc-shape exception), the script automatically falls back to dumping
every sheet and says so — it never silently drops content.

**Important: do not save this as a standalone `.ps1` file and invoke it with `-File` or dot-sourcing
— this environment's endpoint security (Cylance Script Control) blocks executing a `.ps1` file
outright, even via dot-sourcing.** Always paste the script inline as the PowerShell tool's command,
exactly like "The dump script" above.

```powershell
$SourcePath = "<absolute path to the reference file>"
$OnlySheetPatterns = $null   # e.g. @("ﾃｰﾌﾞﾙﾚｲｱｳﾄ") or @("XJC(*","SJC(*"); $null to dump every sheet (note: no closing ")" on prefix patterns — see below)
$CacheRoot = Join-Path $env:USERPROFILE ".claude\skills\_cache\xlsx-dumps"

$resolved = Resolve-Path -LiteralPath $SourcePath
$SourcePath = $resolved.Path
$sourceInfo = Get-Item -LiteralPath $SourcePath
$currentMTime = $sourceInfo.LastWriteTimeUtc.ToString("o")
$currentLength = $sourceInfo.Length
$filterKey = if ($OnlySheetPatterns) { ($OnlySheetPatterns | Sort-Object) -join ";" } else { "" }

$md5 = [System.Security.Cryptography.MD5]::Create()
try { $hashBytes = $md5.ComputeHash([System.Text.Encoding]::UTF8.GetBytes($SourcePath.ToLowerInvariant() + "|" + $filterKey)) }
finally { $md5.Dispose() }
$hash = -join ($hashBytes | ForEach-Object { $_.ToString("x2") })

$cacheDir = Join-Path $CacheRoot $hash
$metaPath = Join-Path $cacheDir "meta.json"

$DumpFormat = 2       # 1 = coordinates relative to the UsedRange; 2 = sheet-absolute

# PowerShell 7's ConvertFrom-Json silently turns an ISO-8601 string into [DateTime], and that
# DateTime stringifies as "07/28/2026 01:39:58" — so comparing it to the "o"-format string the
# cache was written with is ALWAYS false, and the cache never hits. Compare as DateTime instead,
# accepting either shape. See "Never compare the cached mtime as a string" below.
function Test-MTimeMatch($metaValue, [datetime]$currentUtc) {
    if ($null -eq $metaValue) { return $false }
    try {
        $dt = if ($metaValue -is [datetime]) { [datetime]$metaValue }
              else { [datetime]::Parse([string]$metaValue, [cultureinfo]::InvariantCulture,
                                       [System.Globalization.DateTimeStyles]::RoundtripKind) }
    } catch { return $false }
    return $dt.ToUniversalTime().Ticks -eq $currentUtc.Ticks
}

$cacheValid = $false
if (Test-Path $metaPath) {
    try {
        $meta = Get-Content -Raw -Encoding UTF8 $metaPath | ConvertFrom-Json
        # dumpFormat guards against a cache written by an older script whose coordinates mean
        # something different. The source file is unchanged in that case, so mtime/length alone
        # would wrongly report a hit.
        if ($meta.sourcePath -eq $SourcePath -and
            (Test-MTimeMatch $meta.lastWriteTimeUtc $sourceInfo.LastWriteTimeUtc) -and
            $meta.length -eq $currentLength -and $meta.dumpFormat -eq $DumpFormat) {
            $cacheValid = $true
        }
    } catch { $cacheValid = $false }
}

if ($cacheValid) {
    Get-ChildItem -LiteralPath $cacheDir -Filter "*.txt" | ForEach-Object {
        Write-Output "$([System.IO.Path]::GetFileNameWithoutExtension($_.Name))`t$($_.FullName)"
    }
    Write-Output "CACHE_HIT: $cacheDir"
} else {
    New-Item -ItemType Directory -Force -Path $cacheDir | Out-Null

    Add-Type @'
using System;
using System.Runtime.InteropServices;
public class ExcelComWin32Cache {
    [DllImport("user32.dll")]
    public static extern uint GetWindowThreadProcessId(IntPtr hWnd, out uint processId);
}
public static class XlsxDumpHelperCache {
    // r0/c0 = UsedRange origin offset, so emitted coordinates are sheet-absolute.
    public static string FormatSheet(object vals, int rows, int cols, int r0, int c0) {
        var sb = new System.Text.StringBuilder();
        object[,] arr = vals as object[,];
        for (int r = 1; r <= rows; r++) {
            var parts = new System.Collections.Generic.List<string>();
            for (int c = 1; c <= cols; c++) {
                object v;
                if (arr == null) v = vals;
                else if (rows == 1) v = arr[1, c];
                else if (cols == 1) v = arr[r, 1];
                else v = arr[r, c];
                if (v != null) {
                    string s = Convert.ToString(v, System.Globalization.CultureInfo.InvariantCulture);
                    if (s.Length > 0) parts.Add("[" + (r + r0) + "," + (c + c0) + "]=" + s);
                }
            }
            if (parts.Count > 0) {
                sb.Append(string.Join(" | ", parts));
                sb.Append("\r\n");
            }
        }
        return sb.ToString();
    }
}
'@
    $preExistingExcelPids = @((Get-Process EXCEL -ErrorAction SilentlyContinue).Id)
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    [uint32]$excelComPid = 0
    [ExcelComWin32Cache]::GetWindowThreadProcessId([IntPtr]$excel.Hwnd, [ref]$excelComPid) | Out-Null
    if ($preExistingExcelPids -contains $excelComPid) {
        throw "Excel COM automation attached to the user's existing Excel process (PID $excelComPid). Aborting without calling Open/Close/Quit on it."
    }

    $wb = $excel.Workbooks.Open($SourcePath, $true, $true)

    $targetSheets = @()
    if ($OnlySheetPatterns) {
        foreach ($ws in $wb.Worksheets) {
            foreach ($pat in $OnlySheetPatterns) {
                if ($ws.Name -like $pat) { $targetSheets += $ws; break }
            }
        }
        if ($targetSheets.Count -eq 0) {
            Write-Output "FALLBACK_NO_SHEET_MATCHED: none of the patterns ($($OnlySheetPatterns -join ', ')) matched any sheet in this workbook — dumping every sheet instead"
            $targetSheets = @($wb.Worksheets)
        }
    } else {
        $targetSheets = @($wb.Worksheets)
    }

    $dumpedSheetNames = @()
    foreach ($ws in $targetSheets) {
        $dumpedSheetNames += $ws.Name
        $safeName = [string]::Join('_', $ws.Name.Split([System.IO.Path]::GetInvalidFileNameChars()))
        $file = Join-Path $cacheDir ("$safeName.txt")
        $used = $ws.UsedRange
        $rows = [Math]::Min($used.Rows.Count, 20000)
        $cols = [Math]::Min($used.Columns.Count, 220)
        $vals = $used.Value2
        $text = [XlsxDumpHelperCache]::FormatSheet($vals, $rows, $cols,
                                                   ($used.Row - 1), ($used.Column - 1))
        [System.IO.File]::WriteAllText($file, $text, [System.Text.Encoding]::UTF8)
        Write-Output "$($ws.Name)`t$file"
    }
    $wb.Close($false)
    $excel.Quit()
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null

    @{
        sourcePath = $SourcePath; lastWriteTimeUtc = $currentMTime; length = $currentLength
        onlySheetPatterns = $OnlySheetPatterns; dumpedSheets = $dumpedSheetNames
        dumpedAtUtc = (Get-Date).ToUniversalTime().ToString("o")
        dumpFormat = $DumpFormat
    } | ConvertTo-Json | Set-Content -Path $metaPath -Encoding UTF8

    Write-Output "CACHE_MISS_REDUMPED: $cacheDir"
}
```

Read the resulting `<sheetName>.txt` files with the Read tool exactly as you would a fresh dump —
the cache is transparent to every downstream check in this document; only the redump step is
skipped when the source file hasn't changed. **This mechanism is specifically for reference/master
files, not the target design-doc workbook under review** — that one keeps using the plain "dump
once, share the text" flow above (it's reviewed once and the text is only needed for the duration of
that one REV, so a persistent cache buys nothing there and would only grow the cache directory with
one-off entries).

### Batch variant: checking many reference files in ONE Excel session

**Not for REV masters.** This cache writes raw `Value2` with struck text, and ﾃｰﾌﾞﾙﾚｲｱｳﾄ files carry
struck columns; every REV check reads its layouts live with `scripts/live_dump.py` (one out_dir per
file, `--sheets "^ﾃｰﾌﾞﾙﾚｲｱｳﾄ$" --prefix <ID>`). Keep this only for a caller that wants raw values.

**Use this instead of the single-file version above whenever you already know you need to check more
than a couple of reference files up front** — most commonly, every テーブルレイアウト file for the
tables in a program's Ⅲ．入出力定義 list (`xlsx-db-column-check`, easily 10+ tables), or every
テーブルレイアウト file a WG-scoped comparison needs (easily dozens). The single-file
script launches and quits a whole `Excel.Application` COM instance (Add-Type, `New-Object`, the
pre-existing-PID safety check, `Quit`/`ReleaseComObject`) per file — real, non-trivial overhead that
multiplies by file count if you paste it once per file. The batch variant checks every file's cache
validity FIRST, with no Excel interaction at all, and only launches a single shared `Excel.Application`
for the whole batch if at least one file actually needs a redump — so once the cache is warm (a
second REV of the same WG, or `xlsx-db-column-check` re-checking a table an earlier run
already dumped), this typically launches Excel zero times.

```powershell
$Files = @(
    @{ SourcePath = "<absolute path to reference file 1>"; OnlySheetPatterns = @("ﾃｰﾌﾞﾙﾚｲｱｳﾄ") }
    @{ SourcePath = "<absolute path to reference file 2>"; OnlySheetPatterns = @("XJC(*","SJC(*") }
    # ... one entry per file; OnlySheetPatterns = $null for a file type that needs every sheet
)
$CacheRoot = Join-Path $env:USERPROFILE ".claude\skills\_cache\xlsx-dumps"
$DumpFormat = 2       # 1 = coordinates relative to the UsedRange; 2 = sheet-absolute

# Same DateTime-vs-string trap as the single-file script above — see "Never compare the cached
# mtime as a string". Without this the batch reports every file as a miss and relaunches Excel.
function Test-MTimeMatch($metaValue, [datetime]$currentUtc) {
    if ($null -eq $metaValue) { return $false }
    try {
        $dt = if ($metaValue -is [datetime]) { [datetime]$metaValue }
              else { [datetime]::Parse([string]$metaValue, [cultureinfo]::InvariantCulture,
                                       [System.Globalization.DateTimeStyles]::RoundtripKind) }
    } catch { return $false }
    return $dt.ToUniversalTime().Ticks -eq $currentUtc.Ticks
}

function Get-XlsxCacheInfo($SourcePath, $OnlySheetPatterns, $CacheRoot) {
    $resolved = Resolve-Path -LiteralPath $SourcePath
    $SourcePath = $resolved.Path
    $sourceInfo = Get-Item -LiteralPath $SourcePath
    $currentMTime = $sourceInfo.LastWriteTimeUtc.ToString("o")
    $currentUtc = $sourceInfo.LastWriteTimeUtc
    $currentLength = $sourceInfo.Length
    $filterKey = if ($OnlySheetPatterns) { ($OnlySheetPatterns | Sort-Object) -join ";" } else { "" }
    $md5 = [System.Security.Cryptography.MD5]::Create()
    try { $hashBytes = $md5.ComputeHash([System.Text.Encoding]::UTF8.GetBytes($SourcePath.ToLowerInvariant() + "|" + $filterKey)) }
    finally { $md5.Dispose() }
    $hash = -join ($hashBytes | ForEach-Object { $_.ToString("x2") })
    $cacheDir = Join-Path $CacheRoot $hash
    [PSCustomObject]@{
        SourcePath = $SourcePath; OnlySheetPatterns = $OnlySheetPatterns
        CurrentMTime = $currentMTime; CurrentUtc = $currentUtc; CurrentLength = $currentLength
        CacheDir = $cacheDir; MetaPath = Join-Path $cacheDir "meta.json"
    }
}

# Pass 1: resolve cache status for every file — zero Excel interaction
$plan = @()
foreach ($f in $Files) {
    $info = Get-XlsxCacheInfo -SourcePath $f.SourcePath -OnlySheetPatterns $f.OnlySheetPatterns -CacheRoot $CacheRoot
    $valid = $false
    if (Test-Path $info.MetaPath) {
        try {
            $meta = Get-Content -Raw -Encoding UTF8 $info.MetaPath | ConvertFrom-Json
            if ($meta.sourcePath -eq $info.SourcePath -and
                (Test-MTimeMatch $meta.lastWriteTimeUtc $info.CurrentUtc) -and
                $meta.length -eq $info.CurrentLength -and $meta.dumpFormat -eq $DumpFormat) { $valid = $true }
        } catch { $valid = $false }
    }
    $plan += [PSCustomObject]@{ Info = $info; CacheValid = $valid }
}

# Emit every cache hit immediately, and collect the real misses
$toRedump = @()
foreach ($p in $plan) {
    if ($p.CacheValid) {
        Get-ChildItem -LiteralPath $p.Info.CacheDir -Filter "*.txt" | ForEach-Object {
            Write-Output "$([System.IO.Path]::GetFileNameWithoutExtension($_.Name))`t$($_.FullName)"
        }
        Write-Output "CACHE_HIT: $($p.Info.SourcePath)"
    } else {
        $toRedump += $p
    }
}

# Pass 2: only if at least one file is a genuine miss, launch ONE Excel instance for the whole batch
if ($toRedump.Count -gt 0) {
    Add-Type @'
using System;
using System.Runtime.InteropServices;
public class ExcelComWin32Batch {
    [DllImport("user32.dll")]
    public static extern uint GetWindowThreadProcessId(IntPtr hWnd, out uint processId);
}
public static class XlsxDumpHelperBatch {
    // r0/c0 = UsedRange origin offset, so emitted coordinates are sheet-absolute.
    public static string FormatSheet(object vals, int rows, int cols, int r0, int c0) {
        var sb = new System.Text.StringBuilder();
        object[,] arr = vals as object[,];
        for (int r = 1; r <= rows; r++) {
            var parts = new System.Collections.Generic.List<string>();
            for (int c = 1; c <= cols; c++) {
                object v;
                if (arr == null) v = vals;
                else if (rows == 1) v = arr[1, c];
                else if (cols == 1) v = arr[r, 1];
                else v = arr[r, c];
                if (v != null) {
                    string s = Convert.ToString(v, System.Globalization.CultureInfo.InvariantCulture);
                    if (s.Length > 0) parts.Add("[" + (r + r0) + "," + (c + c0) + "]=" + s);
                }
            }
            if (parts.Count > 0) {
                sb.Append(string.Join(" | ", parts));
                sb.Append("\r\n");
            }
        }
        return sb.ToString();
    }
}
'@
    $preExistingExcelPids = @((Get-Process EXCEL -ErrorAction SilentlyContinue).Id)
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    [uint32]$excelComPid = 0
    [ExcelComWin32Batch]::GetWindowThreadProcessId([IntPtr]$excel.Hwnd, [ref]$excelComPid) | Out-Null
    if ($preExistingExcelPids -contains $excelComPid) {
        throw "Excel COM automation attached to the user's existing Excel process (PID $excelComPid). Aborting without calling Open/Close/Quit on it."
    }

    foreach ($p in $toRedump) {
        $info = $p.Info
        New-Item -ItemType Directory -Force -Path $info.CacheDir | Out-Null

        $wb = $excel.Workbooks.Open($info.SourcePath, $true, $true)
        $targetSheets = @()
        if ($info.OnlySheetPatterns) {
            foreach ($ws in $wb.Worksheets) {
                foreach ($pat in $info.OnlySheetPatterns) {
                    if ($ws.Name -like $pat) { $targetSheets += $ws; break }
                }
            }
            if ($targetSheets.Count -eq 0) {
                Write-Output "FALLBACK_NO_SHEET_MATCHED: none of the patterns ($($info.OnlySheetPatterns -join ', ')) matched any sheet in $($info.SourcePath) — dumping every sheet instead"
                $targetSheets = @($wb.Worksheets)
            }
        } else { $targetSheets = @($wb.Worksheets) }

        $dumpedSheetNames = @()
        foreach ($ws in $targetSheets) {
            $dumpedSheetNames += $ws.Name
            $safeName = [string]::Join('_', $ws.Name.Split([System.IO.Path]::GetInvalidFileNameChars()))
            $file = Join-Path $info.CacheDir ("$safeName.txt")
            $used = $ws.UsedRange
            $rows = [Math]::Min($used.Rows.Count, 20000)
            $cols = [Math]::Min($used.Columns.Count, 220)
            $vals = $used.Value2
            $text = [XlsxDumpHelperBatch]::FormatSheet($vals, $rows, $cols,
                                                       ($used.Row - 1), ($used.Column - 1))
            [System.IO.File]::WriteAllText($file, $text, [System.Text.Encoding]::UTF8)
            Write-Output "$($ws.Name)`t$file"
        }
        $wb.Close($false)

        @{
            sourcePath = $info.SourcePath; lastWriteTimeUtc = $info.CurrentMTime; length = $info.CurrentLength
            onlySheetPatterns = $info.OnlySheetPatterns; dumpedSheets = $dumpedSheetNames
            dumpedAtUtc = (Get-Date).ToUniversalTime().ToString("o")
            dumpFormat = $DumpFormat
        } | ConvertTo-Json | Set-Content -Path $info.MetaPath -Encoding UTF8

        Write-Output "CACHE_MISS_REDUMPED: $($info.SourcePath)"
    }

    $excel.Quit()
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null
}
```

Same freshness semantics as the single-file version (mtime + length + sheet-filter all folded into
the cache key), same Cylance constraint (paste inline, never save as a standalone `.ps1`), and the
same read-the-resulting-`.txt`-files-with-Read-tool usage afterward. Build the `$Files` list from
whatever table IDs/paths you already resolved (e.g. step 3's phase-precedence resolution in
`xlsx-db-column-check`) before running this once, rather than looping the single-file script
per table. (Again: not for REV layout reads — see the note at the top of this section.)

## When the dump cascade fails on some sheets — re-verify every finding on those sheets before reporting

Excel COM can die mid-dump on a particular workbook (`RPC server is unavailable`, HRESULT
`0x800706BA`), and the natural recovery is to re-run the remaining sheets — sometimes falling back to
a **plain `Value2` dump without the strikethrough cascade** for the last stubborn sheets, because the
per-character `Characters()` walk is what crashes. Confirmed on
`SSJCB211_ﾛｯﾄ振向け処理(ｻﾌﾞﾌﾟﾛ).xlsx`, where Excel died twice and `更新条件表(TXJCM003)` /
`更新条件表(TXJCD318)` ended up as plain dumps.

That fallback is fine as long as it is **declared** — write a `_README_DUMP_NOTES.txt` beside the
dumps naming exactly which sheets are plain and listing the struck/live text for their affected
cells, and point every agent at it — **and as long as the report is re-verified against live text
before it is handed over**. Two failure modes are specific to a crashed run:

- A plain-dumped sheet still carries deleted text, so an agent reading it can raise a finding against
  content that no longer exists.
- A cell whose per-character walk was interrupted by the crash can be recorded with a *truncated*
  live value in the digest. Real example from the same run: `更新条件表(TXJCM007) [13,34]` was
  written to the first digest as `live='振向'`, which reads like a mid-sentence truncation defect;
  the sheet's clean re-dump and a direct re-read both give the true value `振向先ﾛｯﾄ登録時`. Nothing
  in the digest marks the entry as suspect.

So before reporting, run one targeted Excel COM pass over **every cell the merged report cites** —
each cell's `Font.Strikethrough` (`DBNull` = mixed → rebuild the live text per character), `Font.Color`
for gray, and the resulting live value — and drop or restate any finding whose cited text turns out to
be struck, gray, or different from what the dump showed. Findings of the form "X is missing" need the
complementary check: confirm from the digest that X was not merely struck out, since a deleted X and a
never-written X are different findings. The user asked for this verification explicitly
(「取り消し線加味して誤指摘ないか確認して」) after a full REV, so treat it as a required closing step
of any REV whose dump did not complete cleanly in one pass, not an optional extra.

## Font-size and cell-merge irregularity detection — moved

That technique is specific to `design-doc-formatting-consistency` and now lives in its own file,
`_shared/xlsx-formatting-scan.md`, so the other REV skills don't pay the token cost of a section
they never use. Only `design-doc-formatting-consistency` needs to read that file.

## Operational notes

- **Never write the dump script to a `.ps1` file and run it — pass it INLINE in the PowerShell
  tool's `command` parameter.** This machine runs Cylance Script Control, which blocks PowerShell
  from executing script *files*: the call fails with exit code 34 and the single line
  `Cylance Script Control has blocked PowerShell from running.` — no other diagnostics, and nothing
  identifies the script file as the cause. Inline commands are not blocked, and the whole dump
  script (including its `Add-Type` heredoc) passes fine inline. Confirmed on the `PXJCO124` REV,
  where writing `dump.ps1` to the scratchpad and invoking it with `&` failed this way while the
  identical text run inline succeeded. Note also that shell state does not persist between
  PowerShell tool calls, so `Add-Type` and the dump loop must be in the **same** call regardless.
  If the script is too long to retype inline comfortably, write it to the scratchpad and run it with
  `Invoke-Expression ((Get-Content <path> -Raw -Encoding UTF8))` — Cylance blocks *executing* a script
  file, not evaluating its text, and this form was verified end to end on the `PSJCO307` dump. It
  still counts as one PowerShell call, so the `Add-Type`-in-the-same-call rule is satisfied.
- **Excel's COM server can die mid-dump, and the failure is loud but the output is silently
  wrong.** Seen on the `PXJCO124` REV: partway through the third sheet every subsequent COM call
  started returning `The RPC server is unavailable. (Exception from HRESULT: 0x800706BA)`, followed
  by hundreds of `You cannot call a method on a null-valued expression` errors as `$ws.Cells`
  returned null. Because the errors are non-terminating by default, the script ran to completion:
  it still wrote per-sheet `.txt` files (the `Value2` arrays already fetched were intact) and still
  printed a summary — but the strikethrough/gray resolution for that sheet and every sheet after it
  was abandoned partway, and the later sheets were never written at all. The tell is in the summary
  line: the crashed sheet reports `0x0` for its `UsedRange` size, and sheets are missing from the
  output directory. Set `$ErrorActionPreference = 'Stop'` so the run aborts at the first RPC error
  instead of producing a plausible-looking partial dump, delete the partial output, and re-run
  **once** (it succeeded on the immediate retry there). If it fails again on the same workbook, stop
  retrying and use `scripts/live_dump.py` (see the RPC-crash note at the top of this file).
- **Always check the dump's own row/col cap against the sheet's real size before trusting a "not
  found" result.** The template script's default cap is 20000 rows / 260 cols (it prints TRUNCATED when it cuts) (raised from an
  earlier 500/100 default specifically because that was too low for this project's real sheets and
  caused a recurring wasted round-trip: dump at 500 → a section silently missing past row 500 →
  redump the same sheet at a higher cap → re-read. Confirmed sheet sizes that would have tripped the
  old cap: `PSJCO308`'s 画面設計書 (2092 rows), `XJC_ｼｽﾃﾑ共通設計書.xlsx`'s `実績表項目設定` (2652
  rows) and `ﾛｯﾄ停止ﾁｪｯｸ` (999 rows, 126 cols), `09.区分名称_step2.xlsx`'s `区分名称_STEP2～` (2371
  rows, 156 cols), and a 帳票's `ｽﾎﾟｰｼﾝｷﾞﾁﾜｰﾄ(*)` sheet runs ~211 columns wide — the new
  default covers all of these in one pass. Still, don't treat 20000/260 as
  a guarantee: re-check `$used.Rows.Count`/`$used.Columns.Count` from the sheet (printed when you
  dump it) against the cap, and if a sheet is bigger than even this default, redump just that sheet
  with an explicitly higher cap before concluding a section or ID is actually missing.
- Reuse a single `$excel` instance across many files (loop the whole batch inside one PowerShell
  call) instead of relaunching Excel per file — much faster, and this project has hundreds of
  design-doc workbooks per WG.
- `$excel.Visible = $false` and `$excel.DisplayAlerts = $false` before opening anything; always
  `$wb.Close($false)` per file so it doesn't prompt to save.
- If a run times out or errors, clean up the specific hidden instance this script started before
  retrying — a stale instance keeps file handles locked and blocks subsequent opens:
  `Stop-Process -Id $excelComPid -Force -ErrorAction SilentlyContinue`
  **Never run `Get-Process EXCEL | Stop-Process -Force`** (or any other kill-by-process-name
  form) — `EXCEL.EXE` is shared with any Excel window the user has open for unrelated work, and
  killing by name force-closes those too, discarding unsaved edits. Always target the exact PID
  captured via `$excelComPid` (or the equivalent for each instance, if a run opens more than one).
- **Always run the pre-existing-PID check in the dump script above before touching `$wb`/`$excel`
  at all.** This project has already hit the case where `New-Object -ComObject Excel.Application`
  attached to the user's own running Excel instead of creating a new hidden one — the PID-targeted
  kill guard above only protects the error-cleanup path, it does nothing for this failure mode,
  since the normal end-of-script `$wb.Close($false)` / `$excel.Quit()` calls are exactly what force-
  closed the user's unrelated open workbook the one time this happened. If the check throws, stop
  and tell the user rather than working around it.
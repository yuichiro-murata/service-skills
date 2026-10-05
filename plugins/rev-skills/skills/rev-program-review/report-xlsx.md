# REV 指摘一覧 (Excel) — how to write it

Loaded by `rev-program-review` Step 5 only when the reviewer asks for the workbook.

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

(Colours are Excel COM `.Interior.Color` integers, which are BGR. For openpyxl use the RGB hex: `12419407`=`4F81BD`, `13551615`=`FFC7CE`, `14083324`=`FCE4D6`, `15922414`=`EEF4F2`, `14277593`=`D9DBD9`, `13421823`=`FFCCCC`, `13434879`=`FFFFCC`, `16772300`=`CCECFF`.)

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
(入出力定義 / 内部相互参照 / DBｶﾗﾑ / ID採番 / 更新条件表 / 帳票 / ﾌｧｲﾙ出力 / 記述ﾙｰﾙ / 誤字脱字 / 体裁), and `指摘ID` is
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

**A `Chart.Export` can come back blank — always export one throwaway range after the last band, keep
each band to ≤ ~16 list rows, and size-check every PNG (not only the last).** A tall 33-row band in the
middle of a loop also came back as a 9.6 KB blank on `PSJCO403`; splitting it into 16/17-row bands fixed it. Confirmed twice in a single run on `PSJCO307`: a 4-band export loop wrote
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

**When COM is unavailable** (the user's Excel is open and the PID guard aborts, or COM keeps dying):
build the workbook with openpyxl to the same layout — write 指摘ID and dates as text, set fills,
filter and freeze panes explicitly — read it back with openpyxl to check every cell, and tell the
user the PNG render check was skipped. Never attach to or drive the user's own Excel to render.

---
name: rev-program-review
description: Entry point for a full REV of ONE program's design-doc workbook (機能定義書/画面設計書/ﾁｪｯｸ処理設計書/更新条件表/帳票設計書/ﾌｧｲﾙ出力仕様書). Instead of unconditionally running every check, it first presents the 10 single-program check skills as a checkbox list (AskUserQuestion, multiSelect) so the user picks which checks to run BEFORE any dumping or analysis starts, then runs only the selected ones as one combined pass and reports their findings as a single merged report. Use whenever the user asks to REV/レビュー a design-doc workbook without naming the specific checks they want — in that case do NOT launch the individual check skills directly. Skip the checkbox prompt only when the user already named the checks, or explicitly asked for 全部/all/フルREV.
---

# rev-program-review

Runs a REV of one program's design-doc workbook with a **user-selected scope**. The 10 checks below
are individually expensive (each dumps and re-reads a large workbook, several also read WG-folder
DB-layout files and the 01_Doc common-design workbooks), so running all of them when the user only
wanted two wastes a lot of time and tokens. Ask first, then run only what was selected.

## The 10 single-program check skills

| # | skill | what it checks |
|---|-------|----------------|
| 1 | `design-doc-internal-consistency` | 内部相互参照（ｲﾍﾞﾝﾄ⇔処理概要、ﾚｽﾎﾟﾝｽ/画面項目ID/ﾁｪｯｸID登録、画面3点一致、排他制御）＋処理/ﾁｪｯｸ順⇔ﾎﾞﾀﾝ順、区分名称の09.区分名称突合 |
| 2 | `design-doc-io-table-check` | Ⅲ．入出力定義（CRUD一覧）の双方向網羅性 — 最重量・最高収穫のチェック |
| 3 | `xlsx-db-column-check` | 参照ｶﾗﾑがﾃｰﾌﾞﾙﾚｲｱｳﾄに実在するか＋**画面の桁数とﾃｰﾌﾞﾙﾚｲｱｳﾄの桁数の整合**＋ｺｰﾄﾞ/名称の二重保持ｱﾝﾁﾊﾟﾀｰﾝ(名称がLabelの場合のみ指摘)＋検索条件のTXJAM100保存漏れ |
| 4 | `naming-standard-compliance` | 各ID採番規則・ﾍｯﾀﾞｰ情報・表紙 |
| 5 | `update-condition-completeness` | 更新条件表の網羅性（NOT NULL/主キー/共通項目の設定漏れ、登録日時のUPDATE上書き、更新概要の書式） |
| 6 | `report-design-check` | 帳票設計書（帳票一覧登録、Ⅱ．帳票仕様の記入漏れ、参照先ｴｲﾘｱｽ、取得項目⇔印字項目） |
| 7 | `design-doc-formatting-consistency` | ﾌｫﾝﾄｻｲｽﾞ/ｾﾙ結合のブレ（体裁の衛生） |
| 8 | `design-doc-typo-check` | 日本語の誤字脱字・変換ミス |
| 9 | `file-output-spec-check` | ﾌｧｲﾙ出力仕様書（ﾌｧｲﾙID一貫性、ﾌｧｲﾙ一覧登録、Ⅱ．ﾌｧｲﾙ出力仕様の記入漏れ、共通ﾀﾞｳﾝﾛｰﾄﾞ注釈、画面項目ID番台/登録、ﾍｯﾀﾞｰ⇔明細、明細の参照先、編集方法、ﾌｧｲﾙﾚｲｱｳﾄ一致） |
| 10 | `design-doc-writing-rules` | ｾﾙﾌﾁｪｯｸ記述ﾙｰﾙ（Ⅴ項目定義：大文字変換/ﾎﾞﾀﾝID 9xxx/選択□/No/TextBox桁数・文字種、項番の飛び・重複、取得件数の末行記載、ﾚｽﾎﾟﾝｽ表記、画面名＋画面ID） |

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
options per question, so the 10 checks are split 4 + 3 + 3. Present them exactly like this (labels in
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
- Question 3 — header `記述・体裁`, question「実施する記述ルール・誤字・体裁チェックを選択してください（複数選択可）」
  - 「誤字脱字チェック」— `design-doc-typo-check`
  - 「フォントサイズ・セル結合のブレチェック」— `design-doc-formatting-consistency`
  - 「記述ルールチェック（ｾﾙﾌﾁｪｯｸ観点）」— `design-doc-writing-rules`

All three questions go in **one** `AskUserQuestion` call, so all 10 checkboxes appear in a single
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
and do not fall back to running all 10.

### プロンプトを省略してよいケース

Skip Step 2 and go straight to Step 3 when:

- The user already named the checks ("入出力定義と誤字だけ見て", "誤字チェックして") — run exactly those.
- The user explicitly asked for everything ("全部", "フルREV", "全チェック", "all") — run all 10.
- The user invoked one check skill directly by name — that skill runs standalone; this skill isn't involved.
- `AskUserQuestion` is unavailable (non-interactive / batch / subagent context) — fall back to all 10
  and **state in the report** that the full set was run because the scope couldn't be asked.

Re-ask the scope for each **new** REV request; a selection made for one workbook does not carry over
to the next one unless the user says "同じ観点で" or similar.

## Step 3 — 選択されたチェックの実行

1. **Dump the workbook once, up front** — before launching anything — and hand every check the
   resulting text files. Excel COM is the default; use `_shared/scripts/live_dump.py` instead (validated
   cell-for-cell equivalent) when COM is unavailable or busy, or for a large workbook — on `PXJCO161`
   (9 MB, 8,079 struck cells) COM managed 5 of 58 sheets in 35 min, live_dump all 58 in 580 s. Either
   way, hand every check the
   resulting scratchpad text files instead of letting each one re-dump the same workbook. See
   `_shared/xlsx-excel-com-dump.md`, section "dump once, share the text" — that file ships **inside
   this plugin** (`<plugin root>/skills/_shared/`), not under `~/.claude/skills/`; glob
   `**/rev-skills/**/skills/_shared/xlsx-excel-com-dump.md` if the path doesn't resolve. (Exception:
   `design-doc-formatting-consistency` reads fonts/merges itself (openpyxl preferred, COM as fallback);
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
   file the program's screen-item IDs route to** — the routing table in `_shared/agent-guide.md` maps each ID prefix to
   its `82.画面項目辞書_*.xlsx`, and a program using shared `XJZ`/`SJZ` items needs the `_共通` file
   as well as its own WG's. Subset each to the program's IDs, and pass both paths per file in the
   agent's prompt — the subset to read, the full index to grep.
   In the same step, build the two folder-wide lookups that a check would otherwise rebuild per agent
   with the shipped scripts, cached under `~/.claude/skills/_cache/reference-index/` with the file
   names below. **Compare freshness as values** (datetime / integer), never as strings — Python and
   PowerShell print different fractional-second digits, so a string compare always misses. A cache
   file in an older format (e.g. a `__meta__` key, or a screen list without a `# scope:` line) is
   stale: rebuild it.
   - **区分名称 index** (`group → {区分: 区分名称}` from `09.区分名称_step2.xlsx`) — when
     `design-doc-internal-consistency` is selected (check 11):
     `python _shared/scripts/build_kbn_index.py <...\04_共通設計\09.区分名称_step2.xlsx> idx_kbn-name_09.区分名称_step2.json`
     (~15 s; reads the sheet live; fresh while `source-mtime-utc`/`source-length` match the file).
     It encodes the structure rules — groups separated by a truly blank row (struck rows are not
     separators), sub-headings merged into their parent, titles split on `、`.
   - **Screen-name list** (`画面名 → 画面ID` for every visible 画面設計書 under `01_Doc\08_機能定義書\`, all
     WGs, minus the `agent-guide.md` exclusions) — when `design-doc-writing-rules` is selected (W5):
     `python _shared/scripts/build_screen_list.py <...\01_Doc\08_機能定義書> lookup_screen-name_08_機能定義書_ALL.tsv`
     (4–9 min cold; finds the header row by its `画面ID` label, since row 4 is ﾌﾟﾛｸﾞﾗﾑID in 04_品質管理 /
     05_受注出荷; skips stale copies; writes the sheet-name ID as a 5th column for W5; prints `NON-SCREEN-ID`
     (a 画面ID cell holding a program ID) and `HEADER-MISMATCH` (header 画面ID ≠ the sheet's own ID) lines — doc defects
     worth passing to `naming-standard-compliance` — pass those lines in naming's prompt). Check freshness
     with `build_screen_list.py <root> <tsv> --check` (FRESH/STALE; it counts with its own scope rules and
     checks `# format-version`) — never by recounting the folder yourself.
   - **入力可文字種 table** — when `design-doc-writing-rules` is selected (W1f):
     `python _shared/scripts/build_kind_table.py <...\01_Doc\08_機能定義書\<WG>> lookup_kind_<WG>.tsv` (~45 s;
     freshness by the same `--check` flag). Pass its path; the agent runs
     `kind_consistency.py` on it for the program.
   A check run standalone builds what it needs itself. **Launch order:** the screen-name list can take
   5–9 minutes cold — launch every other check first and launch `design-doc-writing-rules` (and give
   naming the HEADER-MISMATCH / NON-SCREEN-ID lines) when the build finishes, instead of letting an agent
   start against a missing file.
3. Run the selected checks as one combined pass — in parallel background agents when there are
   several. Follow each selected skill's own SKILL.md as the authority for how that check is done;
   this skill only decides *which* checks run. **Each agent's prompt must say**: read
   `_shared/agent-guide.md` (not the whole of `xlsx-excel-com-dump.md`), the dump directory and
   digest path (`_DELETED_DIGEST.txt`, or `_DELETED_DIGEST_<prefix>.txt` from a `--prefix` live dump), the per-sheet `dead=`/`partial=` counts the dump printed, and the
   index/lookup paths that check uses — and that the workbook must never be saved. Do not tell the
   formatting agent to load read-only: its scan needs `merged_cells`, which `read_only=True` lacks. The orchestrator itself
   reads both `agent-guide.md` and `xlsx-excel-com-dump.md`, since it produces the dump.
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
   settles it and takes under a minute. **Apply the same to header-row observations** — a right-hand
   (`3D`) header copy reading `-` was raised 15+ times on the PSJCO501 REV, yet ~115 of 323 PHASE2
   sheets do it; only a copy that *contradicts* the left one is a finding.

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

Reviewers routinely ask for the merged report as a workbook ("Excel一覧に出力して"), so offer it once
at the end of Step 4. **When they accept, read `report-xlsx.md` in this skill's folder** and follow it
exactly — file location (`Downloads`, not the design-doc tree), layout, the A1-notation `セル位置`, the
text-format `指摘ID`, row heights, and render-to-PNG verification. It is kept out of this file so a REV
that never exports does not pay for it.

---
name: design-doc-internal-consistency
description: Review one program's design-doc set (機能定義書/画面設計書/ﾁｪｯｸ処理設計書/更新条件表) for internal cross-reference consistency: events vs the processing overview, response/screen-item/message IDs registered (and no other subsystem's message IDs), Ⅳ message text vs its ID, layout/Ⅴ/Ⅵ agreeing (Hidden items in Ⅵ, item order, ○/〇, yyyy/MM/dd), exclusive control, process/check order vs buttons, 区分名称 and in-text references resolving, pasted pictures and the 画面ｲﾒｰｼﾞ/《参考》画面項目遷移 sheets vs the design, copies edited on one side only, 必須/文字種 checks vs Ⅴ, and sub-program/transition parameters vs the callee. Not the Ⅲ．入出力定義 check nor the self-check writing rules. Use when the user asks to レビュー/REV a 機能定義書 or 画面設計書 or whether it is internally consistent; for a full REV the entry point is `rev-program-review`.
---

# design-doc-internal-consistency

Reviews a single program's design-doc workbook (機能定義書 + 画面設計書 + ﾁｪｯｸ処理設計書 +
更新条件表, usually all sheets inside one `<ProgramID>_<name>.xlsx`) for cross-reference
consistency, grounded in this project's own checklist:
`01_Doc/99.共通資料/設計書記述ルール/05.設計書記述ルール_チェックリスト.xlsx` (sheet
"ﾁｪｯｸﾘｽﾄ"). Read that sheet at the start of a review — it is short (~57 rows) and is the
authoritative list of what "correct" means here; the checks below are the checklist's
programmatically-verifiable items 3-4 through 5-4 (item 3-2/3-3, the Ⅲ．入出力定義 completeness
check, now lives in the separate `design-doc-io-table-check` skill — see below), restated as
concrete diffing steps. The checklist also has purely-manual/subjective items (e.g. "列幅は2、
行幅は14.3か" 1-4) — skip those, they aren't worth automating.

This skill checks **cross-reference consistency** only (does X mentioned here also appear over
there) — see the frontmatter `description` above for exactly which adjacent checks (I/O table
completeness, DB-column existence, ID-numbering, formatting, typos) live in the other nine of the ten single-program checks (the self-check writing rules — Ⅴ per-item rules, 項番 — in
`design-doc-writing-rules`)
instead.

## Environment

**Read `_shared/agent-guide.md` first** — scope selection (`rev-program-review` is the entry point),
the environment, what the live dump has already excluded, how to read the 画面項目辞書 index, and the
reporting conventions common to every check. The masters this check dumps itself
(`04.ﾒｯｾｰｼﾞ管理_*`, `06-09_ﾚｽﾎﾟﾝｽ一覧_*`) are read live with `_shared/scripts/live_dump.py` (never the
COM cache — withdrawn IDs are struck, and the cache keeps them). Because the
dump is live, every row you see is live — which is what makes the "used but not declared" / "missing
from dictionary" diffs below trustworthy without extra work.

**Of the ten single-program checks, this is the one that genuinely needs `_DELETED_DIGEST.txt`.**
Its highest-value findings are asymmetries — one half of a pair edited while the other was left
behind — and the live dump only ever shows you the survivor. Read the digest when a check turns on
what was *removed*:

- a ※-note, a branch table row, or a step number whose counterpart was deleted, leaving a dangling
  definition or a hole in a sequence;
- prose that still advertises an output/behaviour whose implementing rows were deleted (confirmed on
  `SXJCB147`: Ⅰ．機能概要 still said 識別ｶｰﾄﾞ was printed after every 識別ｶｰﾄﾞ output row was struck);
- a "削除" revision note sitting beside content that is *not* struck — genuinely ambiguous, and worth
  reporting as a judgment call rather than resolving yourself.

Conversely, do not report something as "a leftover that should have been deleted" without checking
the digest first: it may already be marked dead, in which case there is nothing to fix. That exact
false finding has been made before (`XJC_ｼｽﾃﾑ共通設計書.xlsx`, `ﾛｯﾄ停止ﾁｪｯｸ`, rows 829-832).

## Checks to run (checklist item → what to diff)

For each check, dump the relevant sheets (per the shared doc) and read them, then compare. Cite the
exact sheet/cell for every finding so it's actionable.

1. **3-4 — Event ↔ processing-overview completeness.**
   The checklist asks only for 「画面設計書に定義されている初期処理およびﾎﾞﾀﾝ押下時のｲﾍﾞﾝﾄ」 (05
   checklist `[19,4]`). Collect the 画面設計書 "Ⅳ．画面項目ｲﾍﾞﾝﾄ詳細" rows whose ｲﾍﾞﾝﾄ内容 (col 13)
   **contains** `起動時` or `ﾎﾞﾀﾝ押下` (`ﾎﾞﾀﾝ押下時` in `PXJCO161` `GXJC161B` `[1975]`-`[1980]`), in every
   screen area, and every step of 機能定義書 "Ⅳ．機能処理概要". Each such
   event should have a traceable step (wording need not match). Flag any with none. `ﾛｽﾄﾌｫｰｶｽ`,
   `ﾀﾞﾌﾞﾙｸﾘｯｸ` and similar events are outside 3-4 — flagging them gave three false positives on
   `PXJCO125` `画面設計書(GXJC125A)` `[415]`/`[417]`/`[418]`. A zoom button (`作業場ｽﾞｰﾑ` `[416]`) is a
   ﾎﾞﾀﾝ押下 event, but process overviews routinely leave zooms out: report a missing one as 低/要確認.
   Pagination buttons (next/prev page, `PXJCO161` `GXJC161A` `[786]`/`[787]`) get the same 低 treatment.

2. **3-5 — Response definitions registered.**
   If 機能定義書 "Ⅶ．ﾚｽﾎﾟﾝｽ定義" defines anything beyond "※ｼｽﾃﾑ共通設計書.ﾚｽﾎﾟﾝｽ 参照", confirm it's
   reflected in the program's own WG registry under `01_Doc/06_システム設計書（一覧、管理台帳）/` —
   `06-09_ﾚｽﾎﾟﾝｽ一覧_<WG名>.xlsx`, one each for `基準情報`/`工程管理`/`品質管理`/`受注出荷` — and fall
   back to `06-09_ﾚｽﾎﾟﾝｽ一覧_共通.xlsx` before calling it unregistered. (The checklist's
   「06-09_基準情報_ﾚｽﾎﾟﾝｽ一覧」 is the 基準情報 template's wording; read it as the program's WG.)

3. **4-12/7-3/8-3 — Screen-item IDs registered in the dictionary.**
   Collect every 画面項目ID from 画面設計書 "Ⅴ．画面項目定義" (rightmost ID column, e.g. `XJC0036`,
   `SJC0662`) — note the ID is usually split across two adjacent cells in the dictionary/master
   files (a 3-letter prefix column, then a 4-digit number column) even though it appears as one
   token in the screen design sheet; concatenate them when extracting from the master.
   **Each ID routes to a dictionary file by its own prefix, not by which WG is under review** —
   `XJZ`/`SJZ` → `_共通`, `XJA`/`SJA` → `_基準情報`, `XJB`/`SJB` → `_受注出荷`, `XJC`/`SJC` →
   `_工程管理`, `XJD`/`SJD` → `_品質管理`, all under `01_Doc/04_共通設計/`; a `MENU…` ID routes on the
   prefix that follows `MENU`. So a 工程管理 program that uses a shared item is checked against the
   `_共通` file for that ID, and checking it against `_工程管理` would report a registered item as
   unregistered. `_shared/reference-index.md` carries the full table and the reason an ID is never
   satisfied by a same-named sheet embedded in another WG's copy. Expect to be handed one index per
   file the program's IDs reach — usually one or two. Flag any ID used in the screen design but
   missing from whichever file should hold it (or vice versa if the dictionary shows an entry whose
   name disagrees with the screen design for the same ID — a rename that wasn't propagated).
   **The dictionary holds the displayed label, so a name matches if it equals either Ⅴ's 画面項目名
   or its `表示` value** (col 11, after dropping a trailing key hint `(ENT)`/`(R)`). Comparing with
   画面項目名 alone gave five false mismatches on `PXJCO125` `画面設計書(GXJC125A)`: `XJZ9002` = `検索`
   against 画面項目名 `検索ﾎﾞﾀﾝ` / 表示 `検索(ENT)` `[439]`; likewise `XJZ9016`, `XJZ9003`, `XJC0760`, and
   `XJC0762` = `関連ﾛｯﾄ/実績削除`, which is the 表示 of the item named `削除区分` `[463]`. Flag only when
   the dictionary name matches neither.
   **Normalise before comparing, and accept containment either way**: strip `ﾎﾞﾀﾝ`, series markers
   `(1-N)`/`(n-m)`, unit parentheticals (`(ｼｰﾄ数)`), a trailing `名`, and `※n`. When 表示 is only `※n`,
   compare with that footnote's text; when one item carries several IDs (a label switched by unit),
   accept a match against any of its names. Without this, `PXJCO161` produced ~90 false mismatches.

   **Don't dump `82.画面項目辞書_<WG名>.xlsx` — read an index of it. See
   `_shared/reference-index.md`.** This check needs only `id → 画面項目名`, and an index of that costs
   one to two orders of magnitude fewer tokens than the dump this step used to take, less again once
   subset to the IDs one program actually cites; the figures and the source state they were measured
   at are in that doc, and are deliberately not repeated here. It also carries the two structural
   traps this step kept hitting: the ID a design doc cites is the `ＩＤ` and `連番` sub-columns
   **joined**
   (`XJC` + `8046` = `XJC8046`), so reading either alone yields IDs that appear in no design doc at
   all; and the dictionary sheets have to be located by their `画面項目ID`/`画面項目名` header rather
   than by sheet name, because every WG's copy names them differently in ways that are not guessable
   from the WG name (`_JAGUR`, `_刷新`, a trailing space).

   The index is built by the orchestrating session before agents launch, not by you. If you were
   pointed at one that is missing or stale, say so rather than building it or opening Excel.

   Work from the index for both directions of this check: an ID cited by the doc but absent from the
   index is unregistered, and a 画面項目名 that differs from the index's is the rename-not-propagated
   finding. Entries retired by strikethrough are excluded from the index, so an ID present in it is
   genuinely registered — but an entry **renamed in place** (old name struck, new name live in the
   same cell) is kept, carrying its live name, so a name mismatch against one of those is a real
   finding, not an artefact.

   You are normally handed **two** paths: a *subset* of the index scoped to this program, to read
   whole, and the *full* index, to grep. The subset is the full index filtered down to the IDs found
   in this program's own dump — so an ID this doc cites that is absent from the subset already **is**
   the unregistered-ID finding. Collect those IDs from cols < 100 only: cols ≥ 100 are the
   revision-memo area, and an ID quoted there is history, not a citation (`PXJCO125` `GXJC125A`
   `[410,107]` `画面項目ID変更(XJC7042→SJZ7002)` pulled the retired `XJC7042` into the subset). Do not dismiss it as "the subset just doesn't cover it". Confirm it
   with one grep of the full index before reporting: absent from both is the finding; present in the
   full index means the subsetting filter missed that ID's form, which is worth saying but is not a
   design defect.

4. **4-8/5-1 — Message IDs registered.**
   Collect every message ID referenced in 画面設計書 "Ⅳ．画面項目ｲﾍﾞﾝﾄ詳細" and in ﾁｪｯｸ処理設計書 (error
   codes). In Ⅳ, scan **both** the ﾒｯｾｰｼﾞ column (col 43) and the 画面項目ID column (col 54): a
   confirm/complete pair is often written as text in ﾒｯｾｰｼﾞ with its IDs in 画面項目ID, in the same
   order (`PXJCO125` `GXJC125A` `[410,43]` `実行前：削除します。よろしいですか？⏎実行後：削除しました。`,
   `[410,54]` `XJZ7001,SJZ7002`) — compare each text with its ID's registered name.
   Error-message IDs are `<3-char JOBｺｰﾄﾞ>-<0/1><5-digit>`, e.g. `XJC-000131`.
   The middle letter of the JOBｺｰﾄﾞ identifies which WG owns that prefix — `JA`=基準情報,
   `JZ`=共通, `JC`=工程管理, `JB`=受注出荷, etc. (extend this mapping as you encounter other WG
   prefixes). Route the lookup by that owning WG, not by which program you're reviewing:
   - **Program's own-WG prefix IDs** (e.g. an `XJC`/`SJC` id inside a 工程管理 program): check
     against that WG's own `01_Doc/04_共通設計/04.ﾒｯｾｰｼﾞ管理_<WG名>.xlsx` (e.g.
     `04.ﾒｯｾｰｼﾞ管理_工程管理.xlsx`). **Read only the program's own prefix sheets, live:**
     `live_dump.py <file> <out> --sheets "^(XJC|SJC)\("` for 工程管理 (open-ended after the `(` —
     sheet names carry a suffix after the prefix) instead of a full-file dump: confirmed for real that
     `04.ﾒｯｾｰｼﾞ管理_工程管理.xlsx` has 11 sheets but only the program's own two prefix sheets are
     ever read from it — a foreign-WG prefix routes to that other WG's own file (see below), never
     to a same-named sheet embedded in this one, so the other 9 sheets in this file are dead weight
     for every 工程管理-program review. (Do not use the COM cross-session cache for this file: it
     keeps struck rows, and withdrawn message IDs are exactly what this lookup must not match.)
   - **`XJZ`/`SJZ` (共通) prefix IDs, from *any* program in *any* WG**: check ONLY against
     `01_Doc/04_共通設計/04.ﾒｯｾｰｼﾞ管理_共通.xlsx` (sheets `XJZ(共通)`/`SJZ(共通)`). **Do not use a
     WG-specific file's own embedded `XJZ(共通)`/`SJZ(共通)` sheet as the source of truth for
     this** — those per-WG copies have been observed to be stale duplicates that drift from the
     real common file (e.g. `04.ﾒｯｾｰｼﾞ管理_工程管理.xlsx`'s `XJZ(共通)` sheet showing an outdated
     wording for `XJZ-000020` while the real `04.ﾒｯｾｰｼﾞ管理_共通.xlsx` has the current, correct
     text that actually matches every design doc). Checking the wrong copy produces a **false
     "text mismatch" finding** — this happened repeatedly across several program reviews before
     being caught. If a WG-local `XJZ(共通)`/`SJZ(共通)` sheet disagrees with the real common
     file, that staleness in the WG-local copy is not itself worth reporting (it isn't the design
     doc under review) — just silently use the common file as ground truth.
   - **Another subsystem's hyphenated ID in a 工程管理 program is NG — 中** (designers' answer
     2026-10-05, 課題193-979; adopted 2026-10-16). When the program ID's subsystem — characters 2-4 of
     the program ID (`PXJCO152` → `XJC`, `PSJAO401` → `SJA`) — is `XJC`/`SJC`
     (`PXJCO128`, `PSJCO208`, `SXJCB141` …), its message IDs may only be `XJC`/`SJC` (either) and the
     共通 `XJZ`/`SJZ`. An `XJA`/`SJA` (基準情報) or any other WG's `<JOBｺｰﾄﾞ>-<6 digits>` is a finding
     **even when its text is right** — `PXJCO128`'s `XJA-000010`「ｺﾋﾟｰ登録対象が選択されていません。」/
     `XJA-000011` were to be replaced by `SJC-000188`/`SJC-000189`. Scan every design sheet, cols < 100
     (the revision memo area, col ≥ 100, quotes old IDs as history). Read the foreign ID's text in its
     owner's `04.ﾒｯｾｰｼﾞ管理_<WG>.xlsx` (live, that prefix's sheets only: `--sheets "^(XJA|SJA)\("` for
     基準情報; locate the columns by header — ﾒｯｾｰｼﾞID col 4, 区分 col 10, ﾒｯｾｰｼﾞ(日本語) col 13 in the
     2026-10 files), then search `04.ﾒｯｾｰｼﾞ管理_工程管理.xlsx`'s `XJC(`/`SJC(` sheets and
     `04.ﾒｯｾｰｼﾞ管理_共通.xlsx` for the same text (normalised as in 4a). Name as the fix every ID with the same
     text **and** the same 区分; one whose text matches but 区分 differs is listed as 要確認. None →
     「工程管理での採番が必要」. When another 工程管理 program already made the same swap (a digest line
     `raw='SJA-000099⏎SJC-000368' live='SJC-000368'`, `PXJCO124` `GXJC124A` `[163,53]`), cite it as support.
     One line per sheet, cells listed. Validated 2026-10-16 on `PXJCO152`: `SJA-000099`
     「ｼｰﾄ「設定」のﾍﾟｰｼﾞ設定の値が不正です。」(B) → `SJC-000368` (same text, B).
     Measured 2026-10-16 on the 122 工程管理 PHASE1-3 books (live ﾁｪｯｸ処理設計書/画面設計書/機能定義書
     text, cols < 100): 9 cells, all in ﾁｪｯｸ処理設計書's error-code column 53 — `PXJCO906` `(GXJC906A)`
     `[57,53]`/`(GXJC906B)` `[105,53]` and `PSJCO404` `(GSJC404B)` `[17,53]` `XJA-100001`, `PSJCOB09`
     `(GSJCB09A)` `[23,53]` `XJA-000011`, `PXJCO152` `(GXJC152A)` `[55,53]`/`[57,53]`/`[231,53]`/`[233,53]`
     `SJA-000099`, `PXJCO193` `(GXJC193A)` `[52,53]` `SJA-000066`. A second validation: `PSJCO404`
     `XJA-100001`「明細件数が{0}件を超えるため、追加できません。」(B, found in the 基準情報 file's
     `XJA(基準_JAGUR)` sheet — a valid source) → `XJZ-100010` (same text once `、` is normalised away, B). The unhyphenated dialog IDs (`SJA8001` …) are not covered by this rule.
   - **Any other cross-WG prefix** — a program whose own subsystem is not `XJC`/`SJC` (a 基準情報
     `PSJAO…` program in the 工程管理 folder: `PSJAO401` `GSJA401B` `[211,53]` cites `XJC-000167`): don't
     assume it's missing — note it separately as "cross-WG reference, couldn't verify without checking
     that other WG's own message file" rather than a confirmed gap.
   - **A short, unhyphenated `<3-char JOBｺｰﾄﾞ><4-digit>` reference (e.g. `XJZ7006`, `XJC7006`,
     `SJC7013`) — the confirm/complete-dialog IDs in 画面設計書 Ⅳ, written in the ﾒｯｾｰｼﾞ column as
     "処理前:"/"完了後:" pairs or in the 画面項目ID column beside `実行前`/`実行後` texts — is a DIFFERENT ID namespace from the hyphenated
     `<JOBｺｰﾄﾞ>-<6-digit>` error-message IDs above, and lives in a different file.** Confirmed for
     real on `PXJCO124_ﾛｯﾄ振向け.xlsx`, `画面設計書(GXJC124A)` row 807 (実行ﾎﾞﾀﾝ: 処理前 `XJZ7006`/
     完了後 `XJC7006`) and row 810 (再印刷ﾎﾞﾀﾝ: `SJC7013`) — none of these exist anywhere in
     `04.ﾒｯｾｰｼﾞ管理_共通.xlsx` or `04.ﾒｯｾｰｼﾞ管理_工程管理.xlsx` (their `000001～`/`100001～`/`200001～`
     numbering blocks don't even reach a `7xxxxx` range), which looks like a registration gap at
     first — but all three IDs turned out to be genuinely registered as `MESSAGE`-category rows
     inside the **screen-item dictionary** files instead: `XJZ7006` in
     `82.画面項目辞書_共通.xlsx`'s `XJZ(共通_JAGUR)` sheet, `XJC7006`/`SJC7013` in
     `82.画面項目辞書_工程管理.xlsx`'s `XJC(工程)`/`SJC(工程)再開発追加分` sheets — each with a
     `カテゴリ`/`localizationid` of `'MESSAGE'` even though the file's overall name says "画面項目".
     When a message-shaped ID doesn't hyphenate and doesn't resolve in the `04.ﾒｯｾｰｼﾞ管理_*` files,
     check the relevant `82.画面項目辞書_*` file's dictionary sheets (the `ＩＤ`+`連番` split, same as
     a normal screen-item ID lookup) before concluding it's unregistered.

   **4a — Ⅳ のﾒｯｾｰｼﾞ文言 ↔ 画面項目ID の一致 (self-check `レビュー観点` No.31 「共通のメッセージと画面項目IDが
   一致しているか」).** Check 4 asks whether a cited ID is registered; 4a asks whether the text written
   beside it is the text registered *for that ID*. Run it on every Ⅳ area (300 of the 358 rows below
   are in `共通`, but G-area events carry confirms too, e.g. `PSJCO309` `GSJC309A` `[1224]`).
   - **Anatomy.** Resolve the ﾒｯｾｰｼﾞ and 画面項目ID columns by header label per area, and read **every**
     column labelled 画面項目ID: `PXJCO152` `GXJC152A` `[497]` has the label at both col 54 and col 58,
     and `[502]` writes `XJZ7006` in 54 and `XJC7017` in 58 — reading only one of them reported a
     missing ID that is not missing (found by validating this rule). Two layouts carry the pairing:
     (1) **event row** (No. is a number): ﾒｯｾｰｼﾞ holds the dialog text, one message per line, usually
     `処理前:…⏎処理後:…`; 画面項目ID holds the IDs, one per line or comma-separated, sometimes labelled
     (`処理前：XJZ7012⏎処理後：SJZ7013`). (2) **footnote/dialog block** under the event table: the event
     row's ﾒｯｾｰｼﾞ is only `※4`, and below it a line whose ID column reads `画面項目ID` opens a block in
     which each message line carries its own ID on the same row (`PXJCO101` `GXJC101C` `[712]`-`[723]`),
     or a dialog table `ﾒｯｾｰｼﾞ | <text> | <ID>` (`PXJCO192` `GXJC192A` `[873]`). The IDs are the
     unhyphenated `MESSAGE` entries of `82.画面項目辞書_*` — of 468 IDs on event rows, 450 are in the
     `7xxx` band, 7 `6xxx` (`SJC6002`, `XJC6001`), 6 基準情報 `8xxx` (`SJA8001`-`8006`) — and only one is a
     hyphenated `04.ﾒｯｾｰｼﾞ管理` ID.
   - **Procedure.** (a) Event rows: split ﾒｯｾｰｼﾞ into message lines; drop blank lines, `※…` note lines,
     bind-variable legends (`{0}:…`), condition-only lines (a whole-line `(…)` such as `(処理前)` /
     `(新規登録の場合)`, or a line ending in `場合`/`時`, optionally followed by `:`/`：` — `PSJCO403`
     `GSJC403A` `[1147,43]` `明細未保存時:` above its message), and lines whose body is `-` or `(ﾒｯｾｰｼﾞなし)`
     (`PXJCO161` `GXJC161B` `[1976]` 処理後:(ﾒｯｾｰｼﾞなし) is a statement, not a missing ID); strip trailing
     `※3,※4` references. A line may start with a label (`処理前:`, `完了後、…場合：`) — compare against
     the whole line **and** every tail after a `:`/`：`, rather than stripping a label by pattern
     (message text itself can contain a colon: `計画年月(計画年月:{0})…`). (b) Footnote rows (No. not a
     number, an ID in the ID column): the text is everything left of the ID column on that row; skip a
     label-only line (`処理後：` `[714]` — its messages follow on the next lines), and read a message
     that wraps onto the next line by joining the following rows up to the next row whose ID column is
     filled (`PXJCO101` `[698]`-`[699]` `『製造ｵｰﾀﾞｰ反映しません。よろしいですか？⏎
     (ｷｬﾝｾﾙを選択した場合は、反映処理を実行します。)』` = `XJC7005`). (c) Look each ID up — unhyphenated
     in the dictionary index (routing as check 3; the `日本語` value is the text), hyphenated in
     `04.ﾒｯｾｰｼﾞ管理_*` (routing above; the message-text column). (d) Normalise both sides: NFKC
     (full/half width, `？`/`?`), remove whitespace, `。`, `、`, `「」`, and turn every bind variable —
     `{0}`, `{{0}}`, `{※1-①}`, `&1` — into one placeholder. On an event row the ID's text must **equal**
     one message line (or a run of consecutive lines — `SJC6002` spans two); on a footnote row it must
     be **contained** in the row's text. Use equality on event rows, never containment: `XJZ9007` = `更新`
     is a substring of `更新します。よろしいですか？`, and containment passed exactly the button-ID findings
     below.
   - **Calibrated 2026-10-05** on the 122 工程管理 `PHASE1`-`PHASE3` workbooks (live text): 358 event rows
     with message text on 147 sheets, plus 38 footnote rows (one label-only). 455 event-row ID/text
     pairs and all 37 footnote pairs (36 contained, 1 wrapped onto the next line) agree; what remains is
     all real:
     - **a button ID where the message ID belongs** (4, 中) — `XJZ9007` (`更新`) beside
       `更新します。よろしいですか？`: `PSJAO205` `GSJA205B` `[83,54]`, `PSJCO801` `GSJC801B` `[149,54]`,
       `PSJCO802` `GSJC802B` `[215,54]`; `XJC9007` (`一括更新`) beside `一括更新します。よろしいですか？`
       `PXJCO129` `GXJC129A` `[504,54]`. Name the fix: `XJC7003` (工程管理) / `XJA7007` (基準情報) for
       更新します。よろしいですか？, `XJC7002` for 一括更新します。よろしいですか？;
     - **registered text differs** (中) — `XJZ7031` is `更新します。よろしいですか? ※更新後、再度ﾛｸﾞｲﾝが必要です。`
       but Ⅳ writes plain `更新します。よろしいですか？` (`PSJCO604` `GSJC604C` `[171,54]`, `PSJCO607`
       `GSJC607A` `[146,54]`): either the note is missing from Ⅳ or the ID is wrong. A difference only in
       repeated punctuation is 低: `SJC7029` `処理中です・・・` vs `処理中です・・・・・` (`PSJCO309`
       `GSJC309A` `[1177,54]`);
     - **a message with no ID** (4, 中) — `-` in 画面項目ID: `PXJCO123` `GXJC123A` `[494]`
       `再印刷します。よろしいですか？` (added by NCRN-7129 without an ID; `SJC7013` is the registered one, as
       `PXJCO152` `[504]` uses), `PSJCO309` `GSJC309A` `[1224]`, `PSJCO403` `GSJC403A` `[1147]`,
       `PSJCO406` `GSJC406B` `[107]` (処理前/処理後 pair, no IDs). Fewer IDs than message lines is the same
       finding for the uncovered line; the corpus has none left, but the self-check sample itself does
       (`GXJA802A` `[275]`: 処理前/処理後 texts, only `処理前：XJZ7029`), so don't read the sample as
       permission;
     - **an ID with no message** (要確認) — ﾒｯｾｰｼﾞ `-` but IDs given: `PSJCO406` `GSJC406A` `[96]` and
       `GSJC406B` `[106]` (`XJC7003 XJC7004`, the 更新 pair, on ﾀﾞｳﾝﾛｰﾄﾞ/画面ｸﾘｱ — likely copied), `PXJCO904`
       `GXJC904A` `[606]` `XJC0252` (= `実績表項目`, the button's own label). The sample's call-out
       「Ⅳ．画面項目ｲﾍﾞﾝﾄ詳細はﾒｯｾｰｼﾞがあるところだけ、画面項目IDが必要」 backs this;
     - **placeholder `XXXXXXX`** (17 cells, 中, one line per sheet) — `PSJCO802` (8), `PSJCOA06` (5),
       `PSJAO205` `GSJA205D` `[69]`, `PSJCO806` `[83]`, `PSJCO403` `GSJC403B` `[120]`, `PSJCO404`
       `GSJC404B` `[121]`;
     - **the ID's dictionary `日本語` is empty** (2, 要確認 — the text cannot be confirmed) — `XJC7042`
       (`PSJCO314` `GSJC314A` `[442]` 処理後 `削除しました。`; it was replaced by `SJZ7002`, see the
       `PXJCO125` `[410,107]` memo) and `SJZ7013` (`PSJCO404` `GSJC404A` `[784]` `一時保存しました。`); both
       dictionary rows are live with a blank `日本語` cell;
     - **a hyphenated `04.ﾒｯｾｰｼﾞ管理` ID on a confirm** (1, 要確認 低) — `PSJCO305` `GSJC305A` `[397]`
       `SJC-000341`: its text matches (`実行します。よろしいですか？`, kind `W`), but every other confirm
       uses the dictionary's `XJZ7006`.
   - **Exemptions.** A ﾒｯｾｰｼﾞ cell holding only `※n` (`PSJCO306` `[561]` `処理前:※1`) — its message and ID
       are in the footnote (`[563]` `SJC6005`), check them there. IDs a message line's `※3,※4,※5` point
       to (`PSJAO601` `GSJA601A` `[1669]` `SJA8003`/`8005`/`8006`) are compared in their footnotes, not
       against the event row. Builders exist for all five dictionaries (`_shared/reference-index.md`);
       treat an ID as unresolvable only if its prefix's index was not built for this run — then it is
       the cross-WG note above, not a mismatch. Order is not checked: label pairing
       (`処理前：`/`処理後：`) is optional and most rows rely on line order. Report one line per row, with
       the Ⅳ cell, the ID, both texts, and the fix when the corpus names one.

   **4b — ﾁｪｯｸ処理設計書 ｴﾗｰﾒｯｾｰｼﾞ ↔ 04.ﾒｯｾｰｼﾞ管理.** Do the same for every ﾁｪｯｸ処理 row: compare the
   ｴﾗｰﾒｯｾｰｼﾞ text (col 40) with the master text of the ID in col 53 (normalised as in 4a, bind variables
   as one placeholder), and the ﾚﾍﾞﾙ (col 31: ｴﾗｰ/警告/情報…) with the master's 区分 (`I`/`W`/`B`/`S`).
   A mismatch is 要確認 (13 on `PXJCO161`). A message cell that only delegates
   (`XJC_ｼｽﾃﾑ共通設計書「ﾛｯﾄ停止」ｼｰﾄ参照`) is not a text — skip it. One cell holding two messages
   (`<警告>…⏎<情報>…`) with one ID is the missing-ID finding for the uncovered message (`PXJCO161`
   `ﾁｪｯｸ処理設計書(GXJC161A)` `[36,40]`/`[36,53]`: only `XJC-100019`; the validation run matched the 情報
   line to `XJC7038`).

5. **4-9/4-13 — Screen layout ↔ item-definition ↔ control-spec three-way match.**
   Collect item names from "Ⅰ．画面ﾚｲｱｳﾄ" (the visual mock), "Ⅴ．画面項目定義" (the item table), and
   "Ⅵ．画面項目制御・出力仕様" (the control matrix). Ⅰ and Ⅴ should list the same set of non-hidden
   items (a Hidden item is not drawn), and Ⅵ should list **every** Ⅴ item, Hidden ones included. Flag any
   item present in one but missing from another. Match areas across
   sections by their `G\d+)` prefix, and merge an area's split Ⅵ tables (8-10 event columns each,
   headers included) before comparing columns between areas.

   **Hidden items belong in Ⅵ** (self-check レビュー観点 No.55「隠し項目（排他ﾌﾗｸﾞなど）が書かれているか」;
   adopted 2026-10-16 — earlier versions skipped them on the 05 checklist's note, which the user
   overruled). A Ⅴ item whose 属性 is `Hidden` — trim spaces/newlines and match the start, since `PXJCO128`
   `GXJC128A` `[591,18]` reads ` Hidden` and some read `Hidden⏎※n`; trim item names the same way when
   matching Ⅴ to Ⅵ (` 詳細ﾘﾝｸ(1-N)`) — (排他ﾌﾗｸﾞ, keys carried to an update) must have a row in its
   area's Ⅵ table, marked 非表示. Report a Hidden item absent from Ⅵ at 低, one line per area naming the
   items; a Hidden item whose Ⅵ row shows it as displayed/usable (○, 表示) is 低 too. Not a finding: a
   Hidden row written `-` in a table where other items change (designers' answer 2026-10-06 No.47 — only
   a table in which *every* item stays unchanged uses `-` for all). The reverse is 要確認: an item Ⅴ
   defines as a displayed Label whose Ⅵ row is 非表示 in every column with no event showing it
   (`PSJCO404` `GSJC404A` `[1846]`-`[1848]` 加工区分/加工分類ｺｰﾄﾞ/状態 — drawn on the 画面ｲﾒｰｼﾞ sheet) — should
   it be Hidden, or is a display event missing? Measured 2026-10-16 (re-counted 2026-10-17 with
   the run-level strike rule): of 1,466 Hidden items in the 122 工程管理 books, 209 never appear in their sheet's Ⅵ text once `※n` is dropped
   from the Ⅴ name (re-counted 2026-10-19; e.g. `PXJCO130` `GXJC130C` `[310]`-`[327]` ﾌｪｰｽﾞID…ﾌｪｰｽﾞ工程ｺｰﾄﾞ) — expect
   findings. Drop `※n` before matching: `PXJCO122` `GXJC122A` Ⅴ `[955]` `処置指示No※1` is Ⅵ `[1137,5]` `処置指示No`
   (非表示) — present.

   **Reading the Ⅰ picture itself (my-review-skills 〔目視〕, adopted 2026-10-21).** Get the pictures with
   `_shared/scripts/extract_images.py <workbook> <scratch>/img --sheets ^画面設計書` (kind `画面ﾚｲｱｳﾄ`; a
   list too wide for one picture continues in strips — read them in anchor-row order, `PXJCO128`
   `GXJC128A` `[7,3]` → `[31,3]` → `[44,2]`). When one sheet draws two screens (`PSJCO201` `GSJC201A`:
   ﾒﾆｭｰﾊﾟﾗﾒｰﾀ="1" at `[9,3]`, ="2" at `[39,3]`), each strip continues the screen whose title cell
   (`[8,3]`/`[38,3]`) precedes it and whose list header it carries on — `[30,4]` belongs to the first,
   `[66,3]` to the second; do not chain strips across screens. Besides the item sets above, compare:
   - **List columns in Ⅴ's order — 低.** The column headers of each drawn list, left to right across the
     strips, against the area's Ⅴ items top to bottom (Hidden and struck rows left out; a header is
     Ⅴ's 表示 or the 画面項目名 without `(1-N)`/`ﾘﾝｸ`/`ﾎﾞﾀﾝ`). One line per area naming the items that
     moved. `GXJC128A` G3 (選択 … 解除理由, 24 columns over three pictures) agrees. Columns drawn as an
     ordered subset of Ⅴ agree when the missing items are shown only in another state (Ⅵ, ※n:
     `GSJC201A` ｼｰﾄ数/作業基板数 switched by 発行単位; one Ⅴ table shared by two drawn screens,
     `GSJC201B` 前工程/後工程) — compare each drawing with the rows visible in its own state. A header
     cut off at the column edge (`識別ｶｰﾄﾞ(共通` for `識別ｶｰﾄﾞ(共通)`) still matches (No.47), and a legend
     drawn above the list (`○：発行済 -：未発行`) is not a column.
   - **Buttons by name and position — 低.** The 共通 buttons drawn (by their caption, `検索(ENT)`) in
     reading order, left to right then top to bottom, against Ⅴ 共通's Button rows by 表示; leave
     閉じる/戻る out as check 9 does. A caption that differs from Ⅴ 表示 is 低 too. Only 共通 buttons are
     ordered here; buttons inside an area (ズーム `Z`, ｶﾚﾝﾀﾞｰ) are covered by the item-set comparison above.
     A button Ⅵ shows in one state only is absent from the other state's drawing by design
     (`GSJC201B` 検査表発行, Ⅵ `[629]` ※6 後工程のみ). Check 9 compares processing order with Ⅴ as a stand-in for the layout; when the picture
     and Ⅴ disagree, say which order the picture shows so the designer fixes the right side.
   - **No column — 要確認, only when the program is inconsistent.** Self-check 観点12 asks for a 明細 No,
     but 41 of the 61 `(1-N)` list areas measured in the 工程管理 Ⅴ tables (2026-10-21) have none, so a
     list without one is not a finding by itself. Report it when another list anywhere in the same
     workbook (any screen, child screens included) draws a No column, naming both: `PXJCO128` `GXJC128A`
     G3 starts with 選択/詳細/対象 while `GXJC128C` G1 `[274]` has No (XJZ0425) — one 要確認 line.
   - **Pagination — 低.** The current page is drawn as `❶` (the self-check workbook
     `XXXXX000_ｾﾙﾌﾁｪｯｸ用.xlsx`, call-out on its 画面設計書 sample, 列37/行25), followed by plain page
     numbers (`< ❶ 2 3 4 5 >`); a plain `1`, or no page control on a list that Ⅳ pages, is a finding —
     one line per list area, naming every picture it appears in (`PSJCO201` `GSJC201B` G3 `[9,3]` and
     `[45,3]`: `< 1 2 3 … 5 >`, Ⅳ `[492]`-`[495]` pages). The control has to appear once per list — a
     continuation strip without it is fine (`GXJC128C` `[22,3]`).
   **Not from the picture:** field colours, `*` and pink required fields, and `XXXXX` lengths are judged
   on the 画面ｲﾒｰｼﾞ sheet by check 14 (EMF rendering shifts colours, designers' answer 2026-10-10 No.15);
   a value cut off at a field's edge is never reported (No.47). Text too small to read is 未確認.
   Budget as in check 13: `PXJCO128` has 9 layout pictures, 11,606 est_tokens.

   **Ⅵ usage conditions (`※n`) against Ⅳ.** A Ⅵ cell holding `※n` instead of `○`/`×` makes the item's
   usability conditional; read the footnote and confirm Ⅳ's event for that item agrees (`PSJCO403`
   `GSJC403A` `[1510]` 自動ｾｯﾄ / `[1512]` 削除ﾎﾞﾀﾝ, `※1` = usable only when the search returns 0 rows —
   reported 要確認 in a v1.20 run). Report a condition Ⅳ contradicts or never mentions as 要確認.

   **Also check the ○/× marker character used throughout "Ⅵ．画面項目制御・出力仕様" against the
   sheet's own legend line** (e.g. "ｲﾍﾞﾝﾄによる項目制御(使用可/不可、編集値) ※"○":使用可､"×":使用不可"
   — one such legend line typically precedes each item-control block). The legend defines the
   "usable" marker as `○` (U+25CB WHITE CIRCLE), but individual matrix cells sometimes instead use
   `〇` (U+3007 IDEOGRAPHIC NUMBER ZERO / kanji "zero") — a different Unicode codepoint that renders
   almost identically and is easy to type by mistake (e.g. via IME kanji conversion of "まる"), so a
   plain visual read of the sheet won't catch it. Confirmed for real on
   `PXJCO152_処置指示発行.xlsx`, sheet `画面設計書(GXJC152A)`: the legend at row 786/803 defines `○`,
   but cells `[649,33]`, `[792,16]`, `[792,28]`, `[793,28]`, `[806,12]`, `[806,28]`, `[806,32]`,
   `[808,12]`/`[808,28]`/`[808,32]` (and likely more throughout the sheet) use `〇` instead. Compare
   each marker cell's character by Unicode codepoint (not just by eye or by a plain text-equality
   check that might normalize them) against the legend's own `○`, and flag every cell using the wrong
   codepoint as a typo to correct to `○`.

   **Also check the letter-casing of any date/datetime format string in "Ⅴ．画面項目定義"'s 表示形式
   column.** This project's screen-side (client) display-format convention is `yyyy/MM/dd` — lowercase
   `y` for year, uppercase `MM` for month, lowercase `dd` for day (the standard Java
   `SimpleDateFormat`-style casing, where case distinguishes month `MM` from minute `mm`). A
   same-looking but wrong-cased variant, most often fully-uppercase `YYYY/MM/DD`, is a real defect,
   not a stylistic choice — confirmed for real on `PSJCO304_着手ﾒｯｾｰｼﾞﾒﾝﾃﾅﾝｽ.xlsx`, sheet
   `画面設計書(GSJC304A)`, rows `353`/`355`/`361`/`363` (TextBox-type 表示開始日/表示終了日 items) and
   rows `422`/`423` (their Label-type display counterparts in the results list) — all six show
   `YYYY/MM/DD` in column 25 (表示形式) where `yyyy/MM/dd` was intended. **Scope this check to
   画面項目定義's 表示形式 column specifically** — a date-format string appearing elsewhere, e.g. inside
   a 更新条件表's 取得内容 column describing how a value is formatted at the DB/SQL layer (e.g.
   `ｼｽﾃﾑ日時(YYYY/MM/DD HH24:MI:SS形式)`), follows Oracle's own `TO_CHAR` format-model convention
   instead (where `HH24`/`MI`/`SS` are themselves Oracle-specific tokens, case-insensitive in that
   context) — don't flag those, they're a different convention entirely and casing carries no meaning
   there. Flag every 表示形式 cell using `YYYY` and/or `DD` (or any other wrongly-cased token relative
   to the `yyyy/MM/dd` convention) as a casing defect to correct. A time part follows the same rule:
   `HH:mm` (minutes lower-case); `HH:MM` is the minute-casing defect (`PXJCO161` `GXJC161E` `[151,26]`
   `YYYY/MM/DD HH:MM`).
   **User decision (2026-10-01): `yyyy/MM/dd` is the rule.** The self-check workbook
   (`01_Doc\08_機能定義書\11_工程管理\XXXXX000_ｾﾙﾌﾁｪｯｸ用.xlsx`, sheet `レビュー観点`) 観点50 says
   「日付がある場合、表示形式は「YYYY/MM/DD」などを記載する」 and its own sample is mixed (`GXJA802B`
   `[156,25]`=`yyyy/MM/dd`, `[164,25]`=`YYYY/MM/DD`); read 観点50 as "write a date format", not as the
   casing to use. Do not re-raise this as a conflict, and keep flagging `YYYY/MM/DD` in 表示形式.
   The *absence* of that annotation is a different matter and belongs to a sibling skill: on a
   更新条件表 sheet a bare `ｼｽﾃﾑ日時` value with no `(YYYY/MM/DD HH24:MI:SS形式)` note is a finding,
   raised by `update-condition-completeness` C8. Casing inside the annotation is never a finding on
   either side.

6. **Screen-item ordering consistency between Ⅳ．画面項目ｲﾍﾞﾝﾄ詳細 and Ⅴ．画面項目定義.**
   Collect the order in which items appear in 画面設計書 "Ⅳ．画面項目ｲﾍﾞﾝﾄ詳細" (each screen-area
   block lists its items' event behavior in a fixed sequence) and, separately, the order they
   appear in "Ⅴ．画面項目定義" (the item table, top-to-bottom, normally following 画面ﾚｲｱｳﾄ/tab
   order). Both sections describe the same set of items and should present them in the same
   relative order within a matching scope (same screen area / same item group) — a reordering
   between the two is a strong signal that one section was edited (an item added, removed, or
   moved) without the other being kept in sync, which often also means the underlying change wasn't
   fully propagated to both places. Flag any pair of items whose relative order differs between the
   two sections (item X appears before item Y in one section but after Y in the other). Exclude
   hidden items from this order comparison (Ⅳ seldom mentions them; their presence in Ⅵ is check 5's
   business). **Ⅳ `選択(ﾍｯﾀﾞｰ)` is Ⅴ `選択`** when the same area of Ⅴ has no `選択(ﾍｯﾀﾞｰ)` row of its
   own: Ⅳ may carry a header select event (marked `(ﾍｯﾀﾞｰ)`), a row select event, or both, while Ⅴ may
   define a single 選択 row (`PSJCO315` `GSJC315A` Ⅳ `[293,5]` `選択(ﾍｯﾀﾞｰ)` only, Ⅴ `[320,5]` `選択` only).
   Match them for the order comparison — when Ⅳ has both, use the position of the first one — and never
   report the name difference or a "Ⅳ item missing from Ⅴ" (user instruction on `PSJCO315`, 2026-10-19,
   via my-review-skills; adopted 2026-10-27). Other findings on those rows are unaffected.
   When you can't tell which of the two orderings is authoritative (neither section
   carries an explicit "this is the source of truth" marker), report the finding as a plain "order
   mismatch" between the two locations without asserting which side is wrong, and let the designer
   decide which one to resequence.

7. **5-4 — Exclusive-control checks present when updating.**
   If the program writes to a table — any 更新条件表 sheet exists, **or** 機能定義書 Ⅲ．入出力定義 has a `○`
   under its `C`/`U`/`D` header for any table (`PSJCO403` `[105,23]`-`[105,26]`; TXJAM100, TSJCD400-402,
   TSJCD404, TXJCM008 at `[106]`-`[111]` — no TSJCD403 row — yet no 更新条件表 sheet) — confirm
   ﾁｪｯｸ処理設計書 mentions a 排他ﾁｪｯｸ, and confirm "Ⅴ．画面項目定義" has a hidden 排他ﾌﾗｸﾞ item. Flag
   if a program updates a table but has neither. Judge this **per screen**: a screen whose events
   update tables with no 排他 while another screen of the same program has it (`PXJCO161` A vs B) is
   要確認.

   **Copied screen sheets.** If two 画面設計書 sheets of the program have identical live-dump text (hash the
   per-sheet dump files; `PSJCO403` `GSJC403C` ≡ `GSJC403D`), report one 要確認 line — a copy-paste that
   was never specialised — rather than letting checks 5/6 repeat each finding on both sheets.

8. **(moved)** The Ⅴ．画面項目定義 per-item writing rules (upper-case note, button-ID band, 選択, No,
   TextBox 桁数/入力可文字種) are now `design-doc-writing-rules` W1.

Checks 9 and 11 also come from the self-check workbook (レビュー観点 No.3/56, No.11/41). They were
calibrated on 2026-10-01 against every visible 機能定義書/画面設計書/ﾁｪｯｸ処理設計書/更新条件表 sheet of
120 工程管理 `PHASE1`-`PHASE3` workbooks (`PXJCO130` excluded — openpyxl cannot open it). Re-derive
every verdict; counts and examples are one snapshot.

9. **処理・ﾁｪｯｸの並びがﾎﾞﾀﾝ順と一致しているか (No.3 / No.56).** The self-check's call-outs: "①は初期処理、
   ②以降はﾎﾞﾀﾝ処理の順番を基本とする" (機能定義書) and "ﾁｪｯｸ処理のｲﾍﾞﾝﾄの順番は画面ﾚｲｱｳﾄのﾎﾞﾀﾝ順と
   合わせる" (ﾁｪｯｸ処理設計書).
   - **Reference order** = the `Button` rows of the screen's Ⅴ．画面項目定義 `共通` area, top to bottom.
   - **機能定義書 Ⅳ．機能処理概要**: section titles `A-①.初期処理`, `A-②.検索処理`, … (col 3); the letter
     is the screen (`A` → `G…A`). Assign each title to the screen by the **block heading** it sits under
     (`E.ﾜｰｸﾌﾛｰ承認ﾌﾛｰ一覧`), not by its own letter — a mislettered title (`F-⑥` inside the E block) belongs
     to E, and its wrong letter is `design-doc-writing-rules` W2's finding. **ﾁｪｯｸ処理設計書(<画面ID>)**: event headings `【検索ﾎﾞﾀﾝ押下時】` (col 3).
   - Match names after stripping `ﾎﾞﾀﾝ` / `押下時` / `押下処理` / `処理` / `【】`. A heading written as a
     sentence still matches when it **contains** a button name (`I-⑥.承認ｸﾞﾙｰﾌﾟ…へ挿入する処理　(<< ﾎﾞﾀﾝ押下時)`
     → `<<`, `I-⑦.承認ｸﾞﾙｰﾌﾟ(ﾕｰｻﾞｰ)行削除処理` → `行削除`). Resolve a contained name in this order: (1) the
     name written immediately before `ﾎﾞﾀﾝ押下` (`(<< ﾎﾞﾀﾝ押下時)` → `<<`, even though `検索` also appears in
     the sentence); (2) otherwise the **longest** button name contained, so `行削除` never also counts as
     `削除`; (3) a tie of equal length is left unmatched. A heading whose **first `【…】`** lists several
     events separated by `、` — one check shared by all of them, live `【新規登録ﾎﾞﾀﾝ、詳細ﾘﾝｸ、流動停止解除ﾎﾞﾀﾝ押下時】`
     — is a merged heading, and the rules above do not apply to it as a whole: split that bracket on
     `、`, match each part on its own, drop only the parts that name a link (the link rule below then
     applies per part, not to the whole heading), and place the heading at the position of the
     **earliest** of its buttons in the reference order (the designer's answer: a merged heading only
     has to sit where its first button is). Its other buttons are not compared. Rule (1) applied to
     the whole heading picks the *last* button and invents an order finding: `PXJCO128`
     `ﾁｪｯｸ処理設計書(GXJC128A)` `[37,3]` sits after ﾀﾞｳﾝﾛｰﾄﾞ and before ｺﾋﾟｰ登録, which is right
     for 新規登録 but was reported as 流動停止解除 out of place (2 false pairs, gone with this rule;
     `PSJCO307` `GSJC307A` `[9,3]` `【検索ﾎﾞﾀﾝ押下時、着手/完了入力ﾎﾞﾀﾝ押下時】` is the other case met, no
     verdict change). Only `、` inside that first bracket counts: a `、` in a sentence after it (`…上記ﾁｪｯｸ
     処理で…場合、実行する`, 13 headings) or in a second bracket naming modes (`PXJAO253` `GXJA253B`
     `【更新ﾎﾞﾀﾝ押下時】【新規登録、ｺﾋﾟｰ登録】`) does not make a merged heading. Do **not** split on `/` —
     button names contain it (`着手/完了入力`); `【実行/集計ﾎﾞﾀﾝ押下時】` (`PXJCO152`) and
     `【実行ﾎﾞﾀﾝ/途中保存ﾎﾞﾀﾝ押下時】` (`PSJCO307` `GSJC307B`) stay under rules (1)-(3). A bracket that
     opens with a check's name and a colon — `【排他ﾁｪｯｸ:更新ﾎﾞﾀﾝ押下時、削除ﾎﾞﾀﾝ押下時】` (`PXJCO163`
     `GXJC163B` `[175,3]`) — is not an event heading but the detail of a check those events share,
     written once after them; leave it out of the order comparison (taken as 更新 it would invent
     更新-after-削除). Judge all of this on the live text: struck parts often undo a merge
     (`PXJCO131` `GXJC131A` `[9,3]`, `PXJCO134` `GXJC134A` `[22,3]`, `PXJCO101` `GXJC101A` `[69,3]`/`[76,3]`
     are single events once the struck `、…` is gone). Never match a heading that names a link or a
     frame/row operation inside an area — `…ﾘﾝｸ押下処理`, `承認ｸﾞﾙｰﾌﾟ枠削除処理`, `通知先行削除処理` — those are
     not 共通 buttons even when a 共通 button's name (`承認`, `削除`) appears in them. Items that match no 共通 button (links, zoom, other-area
     buttons) are left out of the order comparison. But a heading that names a **button** (`…ﾎﾞﾀﾝ押下時`)
     found in **no** area's Ⅴ is a dangling event — a renamed or removed button — report it (中):
     `PXJCO161` `ﾁｪｯｸ処理設計書(GXJC161B)` `[19,3]` `【処理取消ﾎﾞﾀﾝ押下時】`, where the button is now
     途中入力取消. Then flag any pair of matched items whose relative order differs from the reference.
   - **Leave `閉じる` / `戻る` out of the order comparison.** Ⅴ often lists them first (top-right on a
     rich screen) while every process overview puts them last — comparing them generates most of the
     noise (`PSJAO401` `GSJA401B`/`F`).
   - Measured: 9 of 154 機能処理概要 lists and 17 of 155 ﾁｪｯｸ処理 sheets disagree with Ⅴ's order. The
     self-check's own sample (`PXJAO802`) places `停止` third in ﾁｪｯｸ処理 while Ⅴ lists it last, and Ⅴ is
     a proxy for the visual layout rather than the layout itself, so **report order findings as 低**,
     one line per sheet, naming the item that moved (`PXJCO134` `ﾁｪｯｸ処理設計書(GXJC134A)`: 解除 before
     検索/ﾀﾞｳﾝﾛｰﾄﾞ; `PSJCO604` `GSJC604A`/`B`: 製伝修正 and 単価修正 swapped).
   - Separately, **`X-①` must be 初期処理** for a screen: 8 lists start with something else
     (`PXJAO243` `B-①` = 更新, `PSJAO501` `B`-`E`) — 低, unless the screen genuinely has no initial
     processing, which the designer should then say.

10. **(moved)** 項番の飛び・重複 is now `design-doc-writing-rules` W2.

11. **区分名称の参照が 09.区分名称_step2.xlsx と一致するか (No.11 / No.41).** Design docs cite a 区分 by
    group name: `ｼｽﾃﾑ共通設計書.区分名称.仕入先区分 参照`, `区分名称.ﾛｯﾄ停止区分(ﾛｯﾄ停止指示登録) 参照`, and
    occasionally a value: `区分名称.製造ｵｰﾀﾞｰ状態.中断`. Each must resolve in
    `01_Doc\04_共通設計\09.区分名称_step2.xlsx`.
    - **Reading the master**: sheet `区分名称_STEP2～` only (the standing user rule in `agent-guide.md`'s
      "区分名称/区分コード lookups" — `区分名称_～STEP1` is not ground truth). A group is a title
      in col 4 (`承認状態`), then a header row `区分`(col 5) / `区分名称`(col 10) / `ﾘｿｰｽID`(col 57), then
      one row per value. Read an index, not a dump — like the 画面項目辞書. Use the prebuilt index your
      prompt names; standalone, build it with
      `python _shared/scripts/build_kbn_index.py <...\01_Doc\04_共通設計\09.区分名称_step2.xlsx> <out.json>`
      (~15 s, live read; output `{source, source-mtime-utc, source-length, groups: {group: {区分: 区分名称}}}`).
      The script is the authority on structure; in short, it splits a title naming several groups on
      `、` (`ﾛｯﾄｶｰﾄﾞ発行区分、ﾌﾞｰｽｶｰﾄﾞ発行区分、…` row 1011), folds col-4 sub-headings into their parent
      (`合否判定条件` row 120 → `入力ﾀｲﾌﾟ 1:文字の場合` …), and separates groups only on truly blank rows —
      a struck row is not a separator (`製造条件表示区分`: rows 2011-2012 struck, its `入力ﾀｲﾌﾟ 2/4`
      still belong to it). A hand-built index that misses either title shape produced false "missing"
      findings in calibration.
    - **Parsing the reference**: take the text after `区分名称.` — or after `09.区分名称_step2.` when no
      `区分名称.` follows (`PSJCO403` `GSJC403A` `[1401,25]` `09.区分名称_step2.社内加工状態 参照`) — up to `参照`/whitespace; drop an
      unbalanced trailing `)` (it closes an outer `区分名称(ｼｽﾃﾑ共通設計書.区分名称.製造ｵｰﾀﾞｰ状態)`); keep a
      `(qualifier)` as part of the name; resolve `group.value` (value must be one of the group's names) and
      `group.qualifier` → `group(qualifier)` before calling anything missing. **A newline or whitespace
      after `区分名称` is a separator too**, same as `.`: `ｼｽﾃﾑ共通設計書.区分名称⏎実績途中入力区分 参照`
      (`PXJCO161` `GXJC161A` `[842,26]`/`[872,26]`, `GXJC161B` `[2042,26]`, `[2497,26]` and six more in
      col 26) names the group on the next line — reading it as "no group" gave 10-13 false findings.
    - Measured: 525 references, ~441 resolve once the index handles both title shapes. Report:
      - **placeholder `区分名称.XXXX`** — 中 (≈30 on 2026-10-01, `PSJCO403`/`404`/`405`, plus `検索区分(XXXX)` in
        `PSJCOB01`; `PSJCO403` has none as of 2026-10-06); the 区分 was never decided;
      - **group not in the master** — 中. Confirmed: `PXJCO129` `GXJC129B` `[701,26]` `区分名称.ﾛｯﾄ識別`,
        where the item and its 説明 say `ﾛｯﾄ種別` and the master has `ﾛｯﾄ種別` (row 1016) — a wrong name,
        not a missing registration; say the likely intended group when a near name exists. Other
        calibration candidates (`PSJCO801` ｶﾗｰﾊﾟﾀｰﾝ, `PSJCO802` ｲﾝｸ在庫状態/保管状態, `PSJCO402`
        依頼先工場区分) were not individually verified — re-check each against both title shapes above
        before reporting. A group that is truly absent has no ﾘｿｰｽID, so it cannot be displayed
        multilingually;
      - **`区分名称` with no group name at all** — 中: `ｼｽﾃﾑ共通設計書.区分名称 参照` (`PXJCO129` `GXJC129A`
        `[726,26]`, 未作業ﾌﾗｸﾞ, where the same item elsewhere reads `…区分名称.未作業ﾌﾗｸﾞ 参照`). Flag a
        delegation (`…区分名称 参照`, `区分名称(…)`) with no `.<group>`. `区分名称` used as a plain noun is
        not a reference — e.g. the display expression `検索時:(6)①.ﾛｯﾄ構成区分+":"+区分名称` in 説明
        (`PXJCO125` `画面設計書(GXJC125A)` `[478,43]`), whose group `区分名称.ﾛｯﾄ構成区分` is in `[478,25]`.
        Likewise `X状態+":"+区分名称(ｼｽﾃﾑ共通設計書.区分名称)` when `X状態` is itself a group name in the index
        (`PXJCO161` `ﾁｪｯｸ処理設計書(GXJC161B)` `[152,59]`-`[157,59]`, `(GXJC161D)` `[13,59]`-`[20,59]`,
        製造ｵｰﾀﾞｰ状態/移動ﾛｯﾄ状態): the operand names the group — resolved, 低 at most. A bare group name with
        exactly one qualified match in the index (`<group>` → only `<group>(<qualifier>)`) is 要確認 naming
        that candidate;
      - **qualifier missing or different** — 要確認 (11): a bare `検索区分` when the master has
        `検索区分(ﾛｯﾄﾄﾚｰｽ)`/`(不良ﾛｯﾄ統合)`/`(社内加工用)` — say which one is meant; `検索区分(社内加工)` vs the
        master's `(社内加工用)` is a near-miss to correct.
    - **Inline value lists** (`'0':表示しない、'1':表示する`, or an 初期値 like `11:新規発行` / `0:通常`, beside an item named after a group): compare
      code and name with the master when present. Only 7 rows in the corpus carry one and all matched, so
      expect silence.
    - The same scan surfaces `ｼｽﾃﾑ共設計書`(sic — `通` missing) in many references; that is a typo for
      `design-doc-typo-check`, mention it only if that check is not in the run.

12. **【更新条件表(<ID>) <section>】 cross-references resolve.** 機能定義書/画面設計書/ﾁｪｯｸ処理設計書 point
    into 更新条件表 sheets as `【更新条件表(<ID>)】参照` or `【更新条件表(<ID>) <section>】参照`. The sheet
    `更新条件表(<ID>)` must exist (exact ID; trailing spaces in the sheet name don't count), and a section
    must appear in it as a `【…】` heading (col 16) or a `■` heading (col 2) — compare after stripping
    whitespace and an unbalanced `)`. Report each dangling pointer (中), naming the nearest existing
    sheet/section. `PXJCO161` 機能定義書: `[736,26]` `TXJCM999` (no such sheet; `更新条件表(TSJCM999)`
    exists); `[1012,26]`/`[1014,26]` section `処理取消ﾎﾞﾀﾝ押下時` (TXJCD404/TXJCD407 now have
    `【途中入力取消ﾎﾞﾀﾝ押下時】`); `[153,6]`/`[365,7]`/`[371,7]` `流動停止:ﾛｯﾄ停止ﾁｪｯｸ`/`:加工期限ﾁｪｯｸ`
    (TXJCM006 has only `■流動停止`/`【流動停止】`). This check owns these pointers;
    `design-doc-io-table-check` and `xlsx-db-column-check` leave them here (a pointer to a sheet that is
    missing *because the program has no 更新条件表 for a written table* is still reported once there).

    **Other in-text references (adopted 2026-10-17 from my-review-skills 「本文中の参照」「注記の※n」「参照名」, PH2-424…681,
    PH2-457…1365, 課題1614/1621).** Besides 【更新条件表(ID) section】, every pointer written in the text must land on
    something live. Compare after NFKC (apply it to your patterns too: `Ⅱ`→`II`, `ﾎﾞﾀﾝ`→`ボタン`), stripping whitespace and
    a trailing 参照. Measured 2026-10-17 on the 122 工程管理 PHASE1-3 books.
    - **Sheet pointers** `【画面設計書|ﾁｪｯｸ処理設計書|機能定義書|ﾌｧｲﾙ出力仕様書|帳票設計書(ID) …】` and `画面設計書(ID).Ⅲ…(n)` /
      `ﾁｪｯｸ処理設計書(ID).Ⅱ…(n)`: the sheet must exist (a variant sheet `GSJA241B-基本情報`, `帳票設計書(RXJA003)_ﾃﾝﾌﾟﾚｰﾄ未登録時`
      counts by prefix; another program's ID — `GXJC901A`, `PXJCO901` — resolves in that program's workbook: a screen ID
      via the parameter index's `screens`, a program ID via the program-name list (`build_program_list.py`; the index's
      `programs` holds only programs with an argument table; the all-WG screen-name list, when the run built it, covers
      other WGs' screens). An ID that none of these covers is 未確認, not missing). Missing — 中, naming the likely sheet: `PSJAO203`
      機能定義書 `[79,4]` and 17 more `GSJC203A/B/C` for `GSJA203…`; ID typos `GSJC123A` (`PXJCO123` `[12,59]`), `GSJC125A`,
      `GXJC0121A`, `GXJC0129B`, `GSJC906B`, `GXJCO130A`, `SGJC404A`/`GSJCA404B` (`PSJCO404` `[152,4]`/`[162,4]`); placeholder
      `FSJCxxx` (`PSJCO403`/`405` `[75,24]`); `PSJCO402` points eight times at ﾁｪｯｸ処理設計書 sheets the workbook does not
      have. Measured: 50 of 1,204 bracket pointers and 20 of 636 `(ID).Ⅲ(n)` pointers. Accept `の`, `、`, `:` as well as
      `.` between the sheet and the section (`画面設計書(ID)のⅢ…`, `画面設計書(GXJC165A)、Ⅲ.画面表示仕様:(10)明細の保留区分名取得`
      — `PXJCO165` `FXJC013` `[17,4]`/`[177,23]`; the latter names (10) 「明細の保留区分名取得」 while (10) is 「明細ﾃﾞｰﾀ取得」 —
      低). A workbook pointer `【<ProgramID>_<名称>】` belongs to `design-doc-writing-rules` W8 (ID and name); skip it
      here unless W8 is not in the run.
    - **Event and section.** `【ﾁｪｯｸ処理設計書(ID) ○○押下時】` must match a live 【…】 heading of that sheet: a merged heading
      counts when one of its `、`- or `/`-separated parts names the event (designers' answer 2026-10-10 No.84; `PXJCO122`
      【検索/実行ﾎﾞﾀﾝ押下時】), as does a qualifier written elsewhere (`PXJCO192` 「工程ｾｯﾄﾎﾞﾀﾝ(処置指示工程)押下時」 =
      【工程ｾｯﾄﾎﾞﾀﾝ押下時】(処置指示工程)); `No.n` names a check row; `Ⅱ．ﾁｪｯｸ処理詳細.※24.…` resolves to a Ⅱ heading `※24`.
      Dangling — 中 (a typo-only miss 低): `PSJAO241` `[1403,20]` 「Excel取込ﾞﾀﾝ」, `[1534,20]` 「Excel一括出職ﾎﾞﾀﾝ」; `PSJCO309`
      機能定義書 `[140,4]` 入炉帳抽出 (heading 入炉段抽出); `PSJCO404` `GSJC404A` `[1489,30]` 画面ｸﾘｱﾎﾞﾀﾝ押下時 (copied from
      403A, no such event in 404A). 16 found.
    - **Numbers.** `Ⅲ．画面表示仕様(n)` / `(n)を実行` / `(n)で取得` on a 画面設計書 must hit a live Ⅲ `(n)` heading;
      `Ⅱ．ﾁｪｯｸ処理詳細(n)` a live Ⅱ heading (col ≤ 5). **On a ﾁｪｯｸ処理設計書 a bare `(n)` resolves to that sheet's own Ⅱ
      `(n)` first** (`PXJCO124` `GXJC124A` `[161,59]` 「(6)にて」 is Ⅱ `[323,2]` 「(6).帳票ﾏｽﾀ取得」, not the screen's Ⅲ(6)); go to
      the screen's Ⅲ only when `画面設計書(ID).Ⅲ` or `Ⅲ．画面表示仕様` is written. Report a miss (中), and say when the
      number exists only struck: `PSJAO401` `ﾁｪｯｸ処理設計書(GSJA401B)` `[112,59]` → `GSJA401D` Ⅲ(11), which ends at (8).
      `【画面設計書(ID)】Ⅲ．…(n)` resolves in **that** ID's Ⅲ: `PSJAO401` `GSJA401E` `[273,3]` 【画面設計書(GSJA401A)】Ⅲ(3) is
      live at `GSJA401A` `[203,4]` 「(3)Rev制御区分取得」 — not a finding, though GSJA401E's own (3) `[175,4]` is struck. The
      screen-clear wording 「Ⅲ.画面表示仕様.(n)の取得ﾚｺｰﾄﾞ有無による初期値ﾊﾟﾀｰﾝ…」 is
      the sample's own and only its number is checked (designers' answer 2026-10-02 No.80). About 20 found (the
      first count, 25, resolved some pointers in the wrong sheet).
    - **※n, both directions.** Resolve each live `※n` to a live note: a cell at col ≤ 8 starting `※n` with text, a line
      starting `※n` inside any cell, a legend `※2:…、※3:…`, a range `※1～※5`. Notes are numbered per table, so resolve a
      use to the **first definition of that number after it within the same table** (up to the next `1.`/`2.`/`(n)`
      heading — row distance alone ties `PXJCO165` `FXJC013` 明細 `[163,45]` 「日付※4」 to the header's ※4 `[158]` instead of
      the 明細's `[231]`) and, on a 画面設計書, within the same Roman section — a Ⅴ `※3` answered
      only by a Ⅵ legend is the wrong note (`PXJCO129` `GXJC129B` `[669,36]`/`[671,36]` vs `[864,3]`). Report: a marker
      with no note (中; 37 in 28 books); a use whose note is struck (中; `PXJCO161` `ﾁｪｯｸ処理設計書(GXJC161B)` `[184,14]`-
      `[196,14]` 「(※4)」, note `[439,3]` struck); a note nothing references (低; 96). Not uses: `※8桁目以降`-style
      annotations (桁/件/日/文字 after the digits), `※0`, numbers ≥ 50, `※05.ｼｽﾃﾑ共通設計書` (neither a use nor a note —
      `PXJCO165` `[73,5]`/`[78,5]`), and a `※n` inside 【…】 or after
      `(ID).` (it points into another sheet — the bracket rule checks it). A note emptied by strikethrough, and a note whose
      only referrers are struck, are check 16d's — not reported here.
    - **Names.** `(n).X` in Ⅲ/Ⅳ/Ⅴ must be one of Ⅲ(n)'s 取得項目 names (the name given in that table, not the DB column;
      containment either way, NFKC, `(1-N)` and `※n` dropped). Cut X at prose (を/が/に/で/の場合/より/から) and skip an X that
      is the block's own title or ends in 取得/処理/ﾁｪｯｸ/設定; skip (n) blocks without a 取得項目 table (delegations); do not
      let a `No.` header row end the table. 130 mismatches in 30 books, roughly half real: a wrong number is 中 (`PSJCO302`
      `GSJC302A` `[533,43]` `(5).KC品名`, (5) = 検索条件初期値; `PSJCO309` `GSJC309A` `[1408,43]` `(12).連番1`, defined in (13));
      spelling is 低 (`PXJCO101` `GXJC101C` `[838,36]`-`[841,36]` `(2).ｼｰﾄ_個片取数` vs シート個片取数; `PSJAO401` `GSJA401C`
      `[407,35]` 集計ｸﾞﾙｰﾌﾟ名 vs 集計GRP名; `PSJCO502` `GSJC502B` `[941,12]` 仕掛工程ｺｰﾄ). The `<alias>.<x>` form of a
      block-reference alias stays with `xlsx-db-column-check`.
    - **Button names in prose — reading guideline only.** Check 9 owns 【…押下時】 headings. In sentences (機能定義書 Ⅳ, Ⅳ 処理
      内容, 更新条件表 conditions) a `○○ﾎﾞﾀﾝ押下` naming no Ⅴ button of any screen is worth reading after a rename (`PXJCO161`
      機能定義書 `[1012,26]`/`[1014,26]` and `更新条件表(TXJCA317)` `[144,16]` 処理取消ﾎﾞﾀﾝ, now 途中入力取消), but a mechanical
      scan is mostly noise (276 of 3,336 mentions: F9 keys, other screens' buttons, paraphrases) — report only renamed/removed
      buttons you have confirmed, 低.

13. **Pasted pictures ↔ text (images other than Ⅰ．画面ﾚｲｱｳﾄ).** Designers paste tables and figures as
    pictures — ランクチェック表, branch tables, sample files — and those are design content the dump never
    shows: 139 of the 工程管理 PHASE1-3 機能定義書/画面設計書/ﾁｪｯｸ処理設計書/更新条件表/ﾌｧｲﾙ出力仕様書
    sheets carry PNG/JPEG pictures. Run once per workbook, after the live dump (`--hidden`: pictures on
    hidden sheets are in scope too, reported 【非表示】 — `agent-guide.md`):
    `python _shared/scripts/extract_images.py <workbook.xlsx> <scratch>/img --dump <live dump dir> --hidden` →
    `index.tsv` (sheet, kind, anchor range, px, est_tokens, note), one PNG per picture, and `covered.txt`.
    It reads the zip only, so it takes seconds (an openpyxl load of a merge-heavy book such as
    `PSJAO401` takes 10+ minutes — never load the workbook just for pictures).
    - **a. Text hidden under a picture — 中, no image needed; every picture, 画面ﾚｲｱｳﾄ included.** Each
      `covered.txt` line is a live cell lying under a picture (the revision margin, DA列 onward, is
      already left out): on paper it cannot be seen, so it is either a stale sentence the picture
      replaced (remove it) or text the picture hides by mistake (move the picture). Report one line per
      picture, quoting the cells. Met in the 2026-10 corpus run (114 books): `PSJCO315` `GSJC315A`
      picture `[264,4]` covers its own explanation `[264,4]`-`[268,4]` (「工程ｺｰﾄﾞ、ﾌｪｰｽﾞ名…横並びに表示する」
      「例)共通ｺｰﾄﾞﾏｽﾀに添付画像の通り…」); `PSJCO401` `GSJC401A` layout pictures at `[8,3]`/`[34,3]` sit on
      the screen titles 「■A.社内加工単価登録 ※…」; `PSJAO205` `GSJA205C` `[128,37]` and `PSJCO802`
      `GSJC802C` `[117,54]` are stray copies of a screen layout lying over Ⅴ (説明, 画面項目ID cells).
    For b, skip rows whose note is `icon` or `same as …`, and kind `画面ﾚｲｱｳﾄ` (check 5 reads those).
    - **b. Picture content vs the text around it — 要確認.** Read each picture (Read the PNG) together
      with the text rows above/below its anchor and the ※n notes it uses. Report when the picture
      disagrees with live text: a ※n in the picture whose note was renumbered or deleted (`PXJCO152`
      `ﾁｪｯｸ処理設計書(GXJC152A)` picture at `[471,4]` still says `※8` in its S×S cell, while the notes it
      refers to were renumbered `※8`→`※7` at `[463,14]`/`[468,4]` — the digest shows the struck `8`);
      an item name, 区分 value or ○/× verdict that contradicts the text's description or the 取得 logic;
      another program's ID or name left from the copy source. Always 要確認, never 確定 — the picture
      may be the stale side, so name both and let the designer decide. When the text contradicts
      itself and the picture sides with one half, report it anyway, quoting both halves (`GXJC152A`:
      the picture and ⑧ `[590,28]` give 出荷合格「値なし」 its own ○ cells, while `[461,14]`/`[466,14]`
      say 値なし = 不合格 = ｴﾗｰ). If the workbook also carries the same table as live text (a sheet such
      as `表ｲﾒｰｼﾞﾗﾝｸﾁｪｯｸ`, `[6,4]` `※8`), report its stale cells in the same finding.
    - **Not this check's business:** a picture of the old system's spec pasted beside the new text as
      reference — JAGUR IDs (`FMPAP10`, `FDCJM06`, `FXABD001`, `GXAB013`), usually anchored right of the
      form (`GXJC152A` `[769,52]`, `[876,50]`, `[979,54]`). Do not compare it with the text; it is
      history, like the screenshots pasted on 画面ｲﾒｰｼﾞ sheets (whose *cells* check 14 does read).
    - **Budget.** If the summed `est_tokens` of the pictures to read — including the ones you open
      only to see whether they are old-system reference — exceeds ~20% of the workbook's live dump
      (characters / 2, without `_DELETED_DIGEST.txt`), say so before reading them and let the user
      choose: read all, or skip b. (`PXJCO152`: 4 pictures, 2,640 against a 26,970 threshold.) Text you cannot read in a picture is 未確認, not a finding. Cite a picture by sheet and
      anchor cell (`ﾁｪｯｸ処理設計書(GXJC152A) 画像 [471,4]`).

Checks 14 and 15 come from my-review-skills' 「画面ｲﾒｰｼﾞ」/「《参考》画面項目遷移」 sections and the
designers' answers behind them; adopted 2026-10-16, when the shared dump started to include these sheets.

14. **画面ｲﾒｰｼﾞ sheets ↔ 画面設計書.** A `*画面ｲﾒｰｼﾞ*` sheet is the screen drawn in cells — labels, `XXXXX`
    placeholders, `*` on required items, coloured fields — and the 画面設計書's Ⅰ．画面ﾚｲｱｳﾄ picture is pasted
    from it (128 visible ones in the 工程管理 PHASE1-3 books, 2 to 3,163 live cells each).
    - **Which sheets, which screen, which state.** Pair each drawn screen with its 画面設計書 by the title
      above it (`PXJCO128` `画面ｲﾒｰｼﾞﾛｯﾄ停止登録` `[3,2]` `■A.ﾛｯﾄ停止一覧(新ｼｽﾃﾑ)` → screen A →
      `画面設計書(GXJC128A)`), or, without a title, by the labels it shares with Ⅴ. One sheet often draws the
      same screen several times, once per state (`PXJCO128` B: 新規 / ｺﾋﾟｰ登録 / 解除(未解除) / 解除済); first
      work out which mode and 区分 each drawing shows — from its title, its example values (停止区分 `[90,24]`
      「C:品目ｺｰﾄﾞ＋製造ﾛｯﾄNo」) and whether a list is filled — and compare it with **that** column of Ⅵ
      (※n tables included). A drawing whose state you cannot tell is one 要確認 line. Not compared: a drawing
      of a screen that no longer exists — its screen letter has no 画面設計書 sheet and the 表紙 改訂履歴 or the
      digest records its removal (`PXJCO128` `[251,2]` ■D, 表紙 `[47,25]` 「GXJC128Dを除去」) — not reported
      at all; and a sheet whose name marks a leftover (the `agent-guide.md` list, plus `旧…`, `BK_JAGUR…`) —
      one line saying so. Several current variant sheets (`…_事業部説明用`, `…(Web版)`) are each compared.
      When the screen IDs written on the drawing match no sheet, pair by screen name and report the ID
      mismatch here, 低 (`PSJCO404` ■C/■D → `画面設計書(GSJC403D)`/`(GSJC403C)`, whose own headers say PSJCO403).
      **Two versions of one screen on one sheet** (same title, neither marked 旧, one dated: `PSJCO404` ■A
      `[3,3]` and `[135,3]`, `[207,91]` 「2026/10/2版」): decide the current one from the version note and from
      which one Ⅳ/Ⅲ match. If Ⅴ matches neither, do not run a-e item by item — report one 要確認 line: Ⅴ/Ⅵ
      not yet updated to the new drawing (list the items only one side has), and the old drawing to be
      deleted or marked 旧.
    - **Out of scope:** typos (the sheet only has to agree with the design, not read well); cells outside
      the drawn screen, deleted screens and leftovers (designers' answer 2026-10-02 No.19). Borders are not
      in the dump, so take the screen's right edge from the fills: the window frame's gray border
      (`theme:0 tint:-0.25` on `PSJCO404`) and the white background it encloses; failing that, the widest of
      the title band (`C83:BQ85`), the header cells (user name / company label) and the list's last column.
      Cells to the right of the frame are out, except a separate drawing labelled as another area or
      variant (`PSJCO404` `[207,91]`-`[221,…]` 別体型/一体型, for Ⅳ G3-G5) — compare that like a screen. Out, for
      example: `PXJCO128` `[38,183]` `製造ﾛｯﾄNo(部分一致部分一致)`, `[70,255]`, `[71,166]`-`[78,201]`.
      Also out: values that run past a field's width (designers' answer 2026-10-10 No.47 — `X`s cut off at a
      button or the frame show that the 桁数 exceeds the display width, on purpose).
    - **a. X count vs 桁数 — 低.** A placeholder's length (`XXXXXXXXX1XXXXX` = 15; the digits are
      position markers and count) must not **exceed** the item's 桁数; fewer is fine. Report fewer only for
      an item that always has that length (処置指示No = 4). A `code:name` placeholder (`X:XXXXXXXXX1…`,
      `XXXXX:XXXXXXXXX1…`) is two values: compare each part with its own item. Dates and times drawn as
      `YYYY/MM/DD` / `HH:MM` are not counted. A literal prefix drawn as itself (`46A` in `46AXXXXXX1X`)
      counts as characters, and a concatenated placeholder with no separator is compared with the combined
      length Ⅳ/Ⅴ describes (`PSJCO404` `[143,12]` `XXXXXXXXX1X` = 11 = 事業所 3 + 加工部門 8, `[1564,24]`).
      (Designers' answers 2026-10-06 No.88/89; 課題193-959.) The 桁数
      is Ⅴ's for an input item. A Label's Ⅴ 桁数 is `-`: compare it only through its Ⅲ 取得項目 → table column
      length when the run has the layout, else skip it — **never borrow the 桁数 of a same-named input**
      (`PXJCO128` list 停止者 `[67,45]` has 30 X's: it shows (1).停止者 = 作業者名, not G1's 10-digit 停止者 code).
    - **b. Example values vs the current definition — 要確認.** A concrete value drawn on the screen
      (`1:停止`, a 停止区分 name, a 数量単位) must match Ⅴ's 説明/初期値 and the `09.区分名称_step2` group
      of that item (check 11's index) (課題1611). Compare ignoring full/half width and spaces — a difference
      only there (`C:品目ｺｰﾄﾞ＋製造ﾛｯﾄNo` vs the master's `+`) is not reported. The same value must also be
      drawn the same way throughout the sheet — one 工程ｺｰﾄﾞ with two different 工程名, a 振向元/振向先 pair
      naming one code differently — because a wrong example is how a coder misreads where a name comes
      from (低). A value drawn for an item Ⅴ does not have is check 5's missing item; if its label names a
      区分 group, compare the codes with that group anyway (要確認 — `PSJCO404` 社内外区分 `[18,41]` 「1:社内」
      while the master has 1 = 社外).
    - **c. Edit format vs Ⅴ — 低.** The *shape* of a drawn value agrees with Ⅴ's 表示形式/説明
      (`(1).ﾛｯﾄ構成区分+":"+区分名称`) (課題193-974): code+name vs code only (`PSJCO404` `[34,7]`
      `XXXXXXXX：XXXXXXXXX1` while Ⅴ G3 `[1757,43]` shows the code alone), the parts joined, a radio
      group (●/○) drawn where Ⅴ defines CheckBoxes or the reverse. Full/half width of the separator
      (`：` vs `:`) and spacing are not reported.
    - **d. Field colours and `*` — 中 for `*` vs 必須, 低 for colour.** Read the fills with
      `python _shared/scripts/sheet_fills.py <workbook.xlsx> "画面ｲﾒｰｼﾞ" <out.tsv>` (zip only, a second
      or less; one line per filled or non-empty cell: `[r,c]  fill  merge  text`). It does not resolve theme
      colours: `theme:0` = lt1 (white), 1 = dk1, 2 = lt2, 3 = dk2, 4-9 = accent1-6; a tint lightens (+) or
      darkens (−). On `PXJCO128`: labels `0070C0`; input fields `FFFFFF`, `theme:0` or `-` (no fill = white);
      required input pink — `theme:5 tint:0.80` on most fields, `FFDCDC` on one (`PSJCO404`: `FCE4D6`
      throughout); display-only / inactive
      fields gray (`EDEDED`, `A6A6A6`, `theme:0 tint:-0.05`/`-0.15`/`-0.25`/`-0.35`). **Not judged by colour:**
      the background of RadioButtons and CheckBoxes and separators such as `～` (`F9F9F9`); striped list
      rows (white/`EDEDED` alternating); a frameless text-only Label.
      Report, against the state the drawing shows: `*` on an item that is not 必須 in that state, a
      required item with no `*`, or `*` on a Label whose Ⅴ 必須 is not ○ (課題1599; self-check call-out
      画面設計書 列64/行5 — a Label that is 必須 ○ in Ⅴ, such as a file-name field filled by a file dialog,
      `PSJCO404` `[120,4]` *取込ﾌｧｲﾙ / `GSJC403D` `[58,33]`, may carry `*`). A 必須
      written `※n` is evaluated in that state; an either-one-required pair may both carry `*`. A required
      item that changes with a 区分 whose `*` does not change is 要確認 (call-out 列67/行6). An input item
      (Ⅴ 属性 TextBox / ComboBox / …) drawn gray although usable in that state, or a Label drawn as a white
      field (課題208-289), is 低; so is an input drawn white although Ⅵ makes it 非活性 in that state, when
      the same sheet draws inactive inputs gray elsewhere (`PXJCO128` 新規 drawing, 停止区分 C: 管理No
      `[91,24]`, 処置指示No1-3 `[92,24]`…, all × in Ⅵ ※1's C column `[679,16]` — 要確認, it may be drawn
      white on purpose). **Judge colours here, never from the rendered Ⅰ．画面ﾚｲｱｳﾄ PNG** — EMF rendering
      shifts colours (designers' answer 2026-10-10 No.15). A program with no 画面ｲﾒｰｼﾞ sheet gets no colour
      finding; say so in one line.
    - **e. Notes — 要確認.** A judgement written in a note on the sheet (解除済/未解除 and the like)
      matches the 画面設計書's condition; a drawing of a state that needs such a note and lacks it is 低
      (`PXJCO128` `[178,3]` ■B 解除 drawing has no state note, while its sibling `[146,3]` carries 「※まだ解除されていない場合」 —
      `[178,3]` needs the matching 「※既に解除済の場合」).
    Validated 2026-10-16 on `PXJCO128` (one 低/要確認 colour finding) and `PSJCO404` (old/new versions on
    one sheet, three 14b 区分 mismatches, two 14d `*`/必須 mismatches).
    Cite cells as `画面ｲﾒｰｼﾞﾛｯﾄ停止登録 [22,31]`, and say which 画面設計書 item it was compared with.

15. **《参考》画面項目遷移 sheets ↔ 画面設計書.** A `《参考》画面項目遷移*` sheet lays the screen out once per
    operation step (■画面起動 → ■○○入力 → ■実行ﾎﾞﾀﾝ押下 …) with the values each step shows (`PXJCO124`
    has two: `《参考》画面項目遷移` and `…(全層)` — the latter is the G1 全層ﾁｪｯｸ ON state, compared with
    Ⅴ's G5 and its 全層 conditions). It is reference material, so every finding is **要確認**, with a weight
    as in 14 (要確認・中 / 要確認・低); name both sides — either may be the stale one. Its wording is
    proofread by `design-doc-typo-check` as usual.
    - **Reading the steps.** A `…表示` step is the state just after the preceding input's ﾛｽﾄﾌｫｰｶｽ; a
      `…編集` step is after the user overwrote inputs (a changed TextBox/ComboBox there is normal; a
      changed Label must follow Ⅴ's formula); `■○○ﾎﾞﾀﾝ押下` / `…(実行後)` changes only what Ⅳ says that
      event sets or clears, plus Ⅴ's `実行時:` values.
    - **a. X count — 要確認・低** — as 14a, against Ⅴ's 桁数 (`PXJCO124`: 作業者ｺｰﾄﾞ `XXXXXXXXX1XX` = 12
      against Ⅴ G1 No.4 `[906,24]` 桁数 10). An item whose Ⅴ 桁数 is `-` (Labels) is compared only when the
      run has the table layout's 桁数; otherwise skip it.
    - **b. Example values — 要確認・中** — a code drawn with a name must match Ⅴ and `09.区分名称_step2`
      (`PXJCO124` `…(全層)` `[324,114]` 「1:シート」, while 数量単位 is 0=ｼｰﾄ, 1=基板); the same code carries the
      same name within a step and across steps (one 工程ｺｰﾄﾞ drawn 「ＩＰＡ－Ａ」 in one step and
      「1次ｼﾝﾀｰ(850℃ Dry)」 in another); and a lot/layer's fixed attributes (品目ｺｰﾄﾞ, KC品名) stay the same.
    - **c. Edit format — 要確認・低** — as 14c (full-width `シート` vs `ｼｰﾄ`, separators).
    - **d. Each step vs Ⅴ's 初期値・説明 — 要確認・中** — ■画面起動 shows Ⅴ's 初期値 (ﾛｸﾞｲﾝ情報.○○ …): an item
      whose 初期値 is not blank drawn blank there is 要確認・低. A step after an input shows what Ⅴ says that
      event sets (`PXJCO124` 管理No ﾛｽﾄﾌｫｰｶｽ: Ⅴ G4 No.36/37 set 振向先数 = 振向元数 and 残数 = 0, but `[446,53]`
      shows 3,000 / 2,000). A radio/ComboBox default that depends on data (「(7).初期値の値により」) is not a
      finding while the drawn value is one it can take; differing only from the fallback is 要確認・低.
    - **e. Values across steps — 要確認・中** — an item the step did not touch keeps its value (`PXJCO124`
      after 実行: 振向先処置指示No 0000 → 0002, 仕掛工程ｺｰﾄﾞ and 数量単位 changed, 処置指示ﾗﾝｸ blanked, while Ⅳ
      `[818,20]` only sets 振向先管理No); a computed value (振向先数, 残数, 発行枚数 …) changes only on the
      events Ⅴ's formula names.
    - **f. `*` and headings — 要確認・中/低** — match Ⅴ and the current layout, i.e. the 画面ｲﾒｰｼﾞ sheet's
      cells and Ⅴ's item order (Ⅰ is a picture). A step still carrying an old layout's heading row is 中
      — say whether the data under it is drawn in the new columns (`PXJCO124` `[230,*]`, `[383,*]`: headings
      `…KC品名|取引先名|仕掛工程ｺｰﾄﾞ…` over data laid out as `…層No|KC品名|品目ｺｰﾄﾞ…`). `*` on a Label is 低;
      when the same `*` comes from the 画面ｲﾒｰｼﾞ sheet, report both cells in one line with check 14d.
    - **g. Leftovers — 低** — a block away from every step that relates to nothing (a column of serial
      numbers only: `PXJCO124` `[845,3]`-`[883,3]`) (self-check call-out 機能定義書 列57/行53).
    Validated 2026-10-16 on `PXJCO124` (12 findings, none withdrawn on re-reading the cells).

16. **Asymmetric edits — one copy changed, the other left behind; deleted things still used** (my-review-skills
    「同じ処理を複製した箇所…片方だけ直っていないか」 and 「削除した引数・項目・取得処理を、後ろの記載がまだ使っていないか」,
    from the PXJCO124 / SSJCB211 reviews; adopted 2026-10-17). Designers copy a block (G4→G5, 振向元/振向先, 件数取得 vs
    一覧取得 SQL, a screen's ﾁｪｯｸ rows, 本表/WF/履歴 更新条件表) and later fix only one copy, or delete an argument/item/note
    and leave its users live. This is the check that reads `_DELETED_DIGEST.txt` most (see the top of this file).
    - **a. Pair the copies.** Within one workbook and one section (Ⅲ with Ⅲ, Ⅴ with Ⅴ, ﾁｪｯｸ events with ﾁｪｯｸ events), two
      blocks are copies when ≥ 5 and ≥ 60 % of their row keys (first non-numeric cell: item name, ﾁｪｯｸ項目, `A.列`) match, or
      ≥ 4 and ≥ 50 % of whole rows match, after normalising `G\d+)`, `(n)`, `※n` and alias letters (`A1.`→`A.`,
      `結合条件(A1 INNER JOIN A2)`). Align rows by key. ※-notes pair by their wording with `※n` and `G\d+)` removed
      (`PXJCO124` G4 ※3 `[1054,5]` ↔ G5 ※2 `[1213,5]` — different numbers, same note); their continuation rows follow.
    - **b. Struck on one side, still live on the other — 中 when the rows mean the same thing.** For an aligned row pair, take
      the struck fragment of a partially struck cell (digest `PART raw/live`). If the other copy's same column still contains
      that fragment live (and has no strike of its own there), the edit was not propagated: `PXJCO123`
      `ﾁｪｯｸ処理設計書(GXJC123A)` `[973,28]` `C.実績項目IDｺｰﾄﾞ`→`C.実績項目ID` (NCRN-5863「項目名誤りを修正」) while the ｸﾗｽ1 copy
      `[1087,28]` still reads `C.実績項目ｺｰﾄﾞ` (same alias C = TXJCD407); `PXJCO152` `ﾁｪｯｸ処理設計書(GXJC152A)` `[856,28]`/`[970,28]` the same;
      `PXJCO906` `ﾁｪｯｸ処理設計書(GXJC906A)` `[27,59]` moved to 「画面設計書(GXJC906A).Ⅲ.画面表示仕様(3)参照」 while `(GXJC906B)`
      `[74,59]` keeps 「共通項目取得.作業者 参照」; `PXJCO124` `GXJC124A` G4 ※3 `[1054,5]` struck 「読込」 while G5 ※2 `[1213,5]`
      still says 「読込時、…」. Measured 2026-10-17 on the 122 工程管理 books: 1,914 copy pairs, 35 such cells in 16 books;
      about a third are intended differences — read each before reporting, and report anything not clearly the same
      meaning as 要確認.
    - **Not findings (seen in the measurement):** a heading block vs its detail block (`SXJCB147` `RSJC033` (2)ﾍｯﾀﾞｰ `"…"`
      literals as 文字列 vs (3)明細 as 数値 — skip when the other side is a quoted literal); sibling branch blocks (`PSJCO503`
      (3)-① 送品 `B.宛先No` vs (3)-② 納品 `NULL`); copies serving different events or screens (`PSJCO204` 実行 `[88,17]` vs
      再印刷 `[151,17]`; `PSJCO201` A-⑦ vs B-④); a Ⅵ column whose whole area starts in a different state (`PXJCO124` G5 is
      非表示 at 起動時, so `[1571,13]` 非表示 vs G4 `[1443,13]` 表示 is right — skip a column that differs in ≥ 40 % of aligned
      rows); column names of different tables (`PXJCO163` TXJCA317 作業開始日時 vs TXJCM008 作業開始時刻 — the layout decides,
      `xlsx-db-column-check`); fragments that are only 参照/時/場合; a row struck in one copy but moved within its own sheet.
    - **c. Reading guideline (not mechanical).** Where a row of one copy carries a margin memo (col ≥ 100: date +
      変更/修正/追加/NCRN) and **its paired row** in the other copy has no memo of that date (compare row to row — a memo of the
      same date elsewhere in the block does not cover it: `PXJCO124` `ﾁｪｯｸ処理設計書(GXJC124A)` G5 `[120,14]` changed on
      2026/9/16 to 「作業前=1以外、かつ作業中=1以外ならｴﾗｰ」 while G4 `[34,14]` still says 「作業前=1以外はｴﾗｰ」, which makes G4 `[45]`
      unreachable), read the two rows side by side (882 such rows in 38
      books — far too many to list). Likewise a row added to only one copy (612). Report only what is not an intended
      difference, as 要確認, naming both cells: `PXJCO124` `ﾁｪｯｸ処理設計書(GXJC124A)` moved G4's lot checks from 【実行】
      `[79]`-`[85]` (2026/5/15 削除) to 【G4)振向元管理No(1-N)ﾛｽﾄﾌｫｰｶｽ】 `[31]`-`[45]`, while G5 keeps them at 【実行】
      `[119]`-`[133]` and has no ﾛｽﾄﾌｫｰｶｽ block although Ⅴ `[1158,43]` names a G5 ﾛｽﾄﾌｫｰｶｽ event; `PXJCO906` `GXJC906A`
      `[28,14]` 自社内以外の作業者ﾁｪｯｸ has no counterpart in `GXJC906B` 【+(ﾌﾟﾗｽ)ﾎﾞﾀﾝ押下】 `[73]`-`[75]`.
    - **d. Deleted, still used — 中** (when b and d hit the same cell, report it once). (1) An argument/ﾊﾟﾗﾒｰﾀ row whose name cell is wholly struck, with no live row of that
      name in the same table, while live text still uses it: for the program's own 引数定義 (under 「引数のﾁｪｯｸ」/
      「ﾌﾟﾛｸﾞﾗﾑ間の引数定義」) a ﾁｪｯｸ処理設計書 col-3 item equal to the name, `引数.X`/`ﾊﾟﾗﾒｰﾀ.X`, or the English 引数名; for a
      screen's (1)ﾊﾟﾗﾒｰﾀ定義 only `ﾊﾟﾗﾒｰﾀ.X` on sheets of the same screen ID; for a table of arguments passed to a callee only
      `引数.X`. `SSJCB211` 引数29 `処置指示ﾗﾝｸ(1-N)`/`ShochishijiRank` `[88]` struck (NCRN-8996) but `ﾁｪｯｸ処理設計書(SSJCB211)`
      `[30,3]` still requires it; `PXJCO161` `GXJC161B` ﾊﾟﾗﾒｰﾀ 作業者ｺｰﾄﾞ(実績表示) `[178,6]` struck 2026/02/12 but Ⅴ `[2068,36]`
      初期値 is still `ﾊﾟﾗﾒｰﾀ.作業者ｺｰﾄﾞ(実績表示)`. (2) A Ⅴ row struck whole (not live under that name on any screen of the
      book) still referenced as `G#)X`, as a ﾁｪｯｸ処理 col-3 item, or — for a button — as `X時`/`X押下` with X the name
      without `ﾎﾞﾀﾝ` in Ⅳ/Ⅴ text or ※ notes (`PXJCO124` 読込ﾎﾞﾀﾝ `[894]` vs 「読込時、」 `[1213,5]`); ﾁｪｯｸ処理 【…】 headings stay with check 9. (3) A ※-note whose text was struck
      leaving only the marker, still used (`PSJCO305` `ﾁｪｯｸ処理設計書(GSJC305A)` `[51,2]` ※1/`[88,2]` ※2 struck 2026/10/7, used
      at `[22,14]`/`[44,14]`) — but a note whose first line was struck and whose text continues on the next row was
      rewritten, not deleted (`PSJCO204` `FSJC068` `[116,5]`-`[117,5]`). The converse — a note whose only referrer is struck —
      is 低 (`SXJCB147` `RSJC033` ※14 `[956]`, its users `[846,17]`/`[846,23]` struck 2026/7/31). Measured: 136 struck
      arguments → 2 still used; 229 struck Ⅴ items → 2; 7 note cases, 16 orphaned notes. A name shared with a DB/master
      concept is not a use (`PSJCO204` ﾁｪｯｸ `[52,3]` 帳票ﾃﾝﾌﾟﾚｰﾄ is a template-existence check, not the struck Hidden item
      `[855]`) — 要確認 at most.
    - **e. A live row whose margin memo says it was deleted — 要確認.** A row with 「…削除」 in its memo (col ≥ 100) but no
      strikethrough is a forgotten strike or a wrong memo. Report it when rows of the same date next to it *are* struck, or
      when the item has no use left: `PXJCO165` `ﾁｪｯｸ処理設計書(GXJC165A)` 【ﾀﾞｳﾝﾛｰﾄﾞﾎﾞﾀﾝ押下時】 `[36]`/`[37]` live with
      「2026/9/3 益満削除」 at `[36,105]`/`[37,105]` while `[30]`/`[32]`/`[38]`/`[39]` of that date are struck; `画面設計書` `[286]`
      取得項目 作業者ｺｰﾄﾞ live, memo 「2026/6/3 益満（削除）」, unused in G2 Ⅴ. Silent when Ⅴ/Ⅵ still has the live item (the memo
      is merely old: `GXJC124A` Ⅵ `[1325]` 再印刷ﾎﾞﾀﾝ).
    - **Strikethrough detection.** Use the live dump and its digest (`_shared/scripts/xlsx_zip.py` follows the same
      run rule, if you need cells the dump lacks). A cell whose own
      font is struck can still hold live runs that carry their own unstruck formatting — openpyxl and `live_dump.py` keep
      those runs (`PXJCO124` `GXJC124A` `[1021,18]` `Hidden`); a scan that drops the whole cell invents deletions.

Checks 17 and 18 come from my-review-skills' 「必須・文字種ﾁｪｯｸ」 and 「遷移・ｻﾌﾞﾌﾟﾛの引数」 viewpoints (self-check 観点5,
課題1603/1606/1613, PH2-668…1260); adopted 2026-10-17. Counts below are **candidates** measured on the 122 工程管理
PHASE1-3 books — confirm each in the dump before reporting.

17. **Ⅴ 必須 / 入力可文字種 ↔ ﾁｪｯｸ処理設計書.** 「必須欄の記号（○／※n）が、ﾁｪｯｸ処理設計書の必須チェックと一致しているか」,
    「登録・更新画面の入力項目ごとに、Ⅴ の入力可文字種に対応する文字種チェックと、日付項目の日付ﾌｫｰﾏｯﾄﾁｪｯｸが漏れなくあるか」,
    「チェックの前提条件が、画面の必須条件・制御表と一致しているか」. Pair each `画面設計書(<ID>)` with every
    `ﾁｪｯｸ処理設計書(<ID>…)` of the same screen ID (a suffix such as `_共通` counts); a screen with no check sheet is out.
    - **Reading the check sheet.** Locate columns by the header row (`ﾁｪｯｸ項目`, `区分`, `ﾁｪｯｸ内容・経緯`, `ﾚﾍﾞﾙ`, `出力先`,
      `ｴﾗｰﾒｯｾｰｼﾞ`, `ｴﾗｰｺｰﾄﾞ`, `ﾁｪｯｸ詳細`, `ﾁｪｯｸ補足` — cols 1 (value in 3)/11/14/31/34/40/53/59/86 on 187 of 189 sheets; resolve
      by label). `【…押下時】` rows in col 3 open an event. A row with col 3 empty inherits the previous ﾁｪｯｸ項目 (`PXJCO128`
      `ﾁｪｯｸ処理設計書(GXJC128B)` `[11]`-`[13]` = 管理No: 未入力, 文字種, 存在 — designers' answer 2026-10-10 No.31). Some sheets
      put a group in col 3 and the item in col 6 (`PXJAO701` `GXJA701B` `(資源情報)|事業所ｺｰﾄﾞ`). Conditions live in 14
      (`※…の場合`), 59 or 86 (`[11,86]` 「画面. 停止区分が管理Noの場合ﾁｪｯｸ」), or in col 3 from its second line on (`PXJCO124` `GXJC124A` `[98,3]`
      「※…場合のみ」, `[115,3]` 「※選択された明細のみ」). A row whose 区分 (col 11) reads
      `…ﾁｪｯｸ処理設計書(<ID>)_共通.Ⅰ.【<item>】 参照` delegates that item (`PSJCO509` `GSJC509A` `[10]`/`[12]`/`[13]`): check it in
      the cited sheet, never report it here.
    - **Matching a check row to a Ⅴ item.** Normalise both names (NFKC, whitespace, a `G\d+)`/`画面.` prefix, `(FROM)`/`(TO)`,
      `(1-N)`, `※n`, `ﾎﾞﾀﾝ`); split a ﾁｪｯｸ項目 on `⏎`, `、`, `+`/`＋` and match each part; a `G2).` prefix restricts the match to
      that area; `X(Y)` in Ⅴ also answers to `Y` when both start alike (`着手日時(着手日)` ↔ `着手日`). When an input item and a
      Label share a name (search `発送日(From)` vs list `発送日`), the check belongs to the input.
    - **Recognisers** (match labels and texts after NFKC). 必須: a row whose error code/message is a required one —
      `XJZ-000001`, 「入力してください。」, 「選択してください。」, 「単位が…入力してください」 — or whose ﾁｪｯｸ内容 *is* a 未入力/未選択/必須
      check; a 未入力 that only states a precondition of another check (`PXJCO165` `[22]`/`[36]`, a range check XJZ-000018
      「…(FROM)と(TO)が未入力で…」) does not count. Not 必須: row-selection errors (`XJZ-000030`/`000023`/`000034`/`000010`,
      「…選択されていません」, 「1件も選択されていない場合ｴﾗｰ」 — `PXJCO124` `GXJC124A` `[167]`); CheckBox `選択(1-N)` items are out.
      When one ﾁｪｯｸ項目 cell names several items (`作業終了日(FROM)⏎作業終了時刻(FROM)`, `PXJCO165` `[20,3]`), assign the check
      only to the item that is the subject of 未入力 in the text (日; the 時刻 is the trigger). Keep `(FROM)` and `(TO)`
      apart when both exist — a FROM check does not cover TO.
      文字種: `文字種|数値ﾁｪｯｸ|半角ﾁｪｯｸ|数値形式` or 「…以外は使用できません」/「数値を入力してください」 (`XJZ-000002`, old form).
      日付: a row whose text says `日付ﾌｫｰﾏｯﾄ` (XJZ-000018 alone is not enough — FROM>TO range checks use it too,
      `PXJCO165` `[16]`/`[33]`). Expected message per 入力可文字種: `文字列(半英数)` → `XJZ-000017`, `文字列(半英数記号)` →
      `XJZ-000016`, `文字列(半数)` → `XJZ-000054`, `数値(整数)正数` → `XJZ-000051`, `数値(整数)` → `XJZ-000050`,
      `数値(小数)正数` → `XJZ-000052`.
    - **a. Ⅴ 必須 ○/※n without a 必須ﾁｪｯｸ — 中** (TextBox/TextArea). A ComboBox/RadioButton needs one only when its 初期値 is
      `空白`/`-` and 説明 does not say `※空白なし` — `PXJCO128` `GXJC128B` `[521]` 停止区分 (初期値 `新規:"C:…"`) and every
      RadioButton in the corpus have a default and cannot be empty. Confirmed: `PSJCO314` `GSJC314A` `[478]` 翌月分 ○, while
      【生産予定数計算ﾎﾞﾀﾝ押下時】 checks only 当月出荷予定数 `[32]`. A repeating-row item (`PSJCO309` `GSJC309A` `[1462]` 管理No ○,
      only existence checks at `[40]`-`[45]`/`[87]`-`[99]`) and ○ on several search conditions with no check at all (`PSJCO201`
      `GSJC201B` `[528]`-`[543]`) are 要確認 — "blank rows allowed" / "one of them required" may be meant.
    - **b. 必須ﾁｪｯｸ without Ⅴ 必須 — 中 when the check is unconditional, 低 when conditional** (Ⅴ should then carry `※n`):
      `PSJCO306` `GSJC306A` `[651,34]` 資源ｺｰﾄﾞ `-` vs `[15]` 「入力必須ﾁｪｯｸ」; `PSJCO201` `GSJC201B` `[582,33]` vs `[26]`. Treat as
      conditional a check whose text (after removing the stock phrase 「未入力の場合はｴﾗｰ」/「入力必須ﾁｪｯｸ」) or ﾁｪｯｸ補足 contains
      `場合|時|のみ|件も|いずれか|どちらか|両方|全て`, or that sits under a `…変更時`/`ﾛｽﾄﾌｫｰｶｽ` event. The inverse pairs — Ⅴ `※n` but
      every check unconditional, Ⅴ ○ but every check conditional — are 要確認: read the ※ footnote and the check's condition side
      by side. A condition expressed by the event heading itself is consistent (`GXJC128B` 解除理由 `○※7`, ※7 = 解除時, checked
      under 【解除ﾎﾞﾀﾝ押下時】), and so is a ○ item whose 属性 switches by mode.
    - **c. 文字種ﾁｪｯｸ / 日付ﾌｫｰﾏｯﾄﾁｪｯｸ coverage — 低, 中 when siblings on the same sheet have it.** Every TextBox/TextArea whose
      入力可文字種 is not `文字列`, `ﾊﾟｽﾜｰﾄﾞ`, `-` or `※n` needs a 文字種ﾁｪｯｸ; `年月日`/`年月`/`月日` types need a 日付ﾌｫｰﾏｯﾄﾁｪｯｸ.
      `時間(4桁)` is exempt (no screen in the corpus writes a time-format check). Judge **per event that submits the item**
      (検索, 実行, ﾀﾞｳﾝﾛｰﾄﾞ …): an event that uses the input without the check is the finding, even when another event has it
      — 要確認 when the check was deliberately removed from that event (`PXJCO165` 【ﾀﾞｳﾝﾛｰﾄﾞﾎﾞﾀﾝ押下時】 lost its 日付ﾌｫｰﾏｯﾄﾁｪｯｸ
      rows `[30]`/`[32]`/`[38]`/`[39]` on 2026/9/3 while 【検索ﾎﾞﾀﾝ押下時】 keeps `[13]`/`[15]`/`[18]`/`[19]`). Also confirmed:
      `PXJCO124` `GXJC124A` 識別ｶｰﾄﾞ発行枚数(1-N) G4 `[1043]` / G5 `[1205]` (TextBox `数値(整数)正数`) have no check row at all,
      while their sibling 振向先数 has XJZ-000051 at `[91]`/`[136]` — 中. Confirmed: `PXJCO129` `GXJC129B` `[617]`-`[644]`
      (仕掛工程ｺｰﾄﾞ…依頼先ｺｰﾄﾞ, `文字列(半英数)`) have 入力必須 + ﾏｽﾀ存在 only, while 事業所ｺｰﾄﾞ on the same sheet has `[25]` 文字種ﾁｪｯｸ
      XJZ-000017. **Search conditions are excluded** (self-check call-out ﾁｪｯｸ処理設計書 「一覧画面の検索条件は、日付・数値以外の
      文字種チェックを書かない」): in a `…検索条件領域` area only dates and numbers need a check, and a string-type 文字種ﾁｪｯｸ there
      is itself a 低 finding (23 rows in 10 books against 334 that follow the rule; `PXJCO130` `GXJC130A` `[236]` 管理No ↔ `[14]`).
    - **d. 文字種 message vs Ⅴ 入力可文字種 — 要確認・中** (課題1603; one side is wrong): `PXJCO101` `GXJC101A` `[725,31]`
      `文字列(半英数)` vs `[60]` XJZ-000016 「半角英数記号以外…」; `PSJCO510` `GSJC510A` `[114,31]` `数値(整数)` vs `[10]` XJZ-000051.
    - **e. Checks written for an item that cannot be entered — 要確認.** A 必須/文字種/日付 check whose only match is a Label,
      Hidden or ReadOnly Ⅴ row. Not findings: an 属性 carrying `※n` (`PXJCO163` `GXJC163B` `[952]` `ReadOnly⏎※2` 工程ｺｰﾄﾞ — an
      input in 新規); a Label filled by a zoom or dialog and marked 必須 ○ (`PSJAO704` `GSJA704J` `[256]` 承認ﾌﾛｰ名); a state check
      on a displayed value (`…がNULLの場合ｴﾗｰ`, `0件`); a group check (`1件も入力されていない`); unit-switched quantity sets.
      A Label row with TAB/必須/入力可文字種 filled is `design-doc-writing-rules` W1e's, not this check's.
    Validated 2026-10-17 on `PXJCO124` and `PXJCO165` (not used in drafting). Report one line per sheet and kind, Ⅴ cell
    and check cell side by side. Measured 2026-10-17 (162 screen/check pairs,
    2,069 input items): a 22 TextBox candidates (ComboBox 17 → 3 after the 初期値 rule), b 73 unconditional / 155
    conditional, c 228 文字種 + 36 日付, d 5, e 56 mostly excluded by the shapes above. Expect a third to half of the
    raw candidates to fall to the exclusions — read each one.

18. **Parameters across workbooks — sub-program calls and screen transitions** (self-check 観点5 「サブプログラムの設計書の
    引数と一致しているか」; 「受け取るﾊﾟﾗﾒｰﾀが、遷移元が渡すﾊﾟﾗﾒｰﾀ（項目・順番・値の出所・型）と一致しているか」, 課題1606). The
    callee is usually in another workbook: read only its definition, from the **parameter index** your prompt names (built by
    the orchestrator with `python _shared/scripts/build_param_index.py <...\01_Doc\08_機能定義書\<WG>> <out.json>` — program →
    引数/戻り値 rows, screen → Ⅱ (1)ﾊﾟﾗﾒｰﾀ定義 rows, ~50 s for 122 books, `--check` for freshness). Standalone, build it yourself.
    Then open the callee's cell in its workbook only to quote it. A callee outside the index (`GXJZ…`, `GSJA…` screens,
    `SXJCB184`) is one line 「呼び出し先の設計書が未入手のため未確認」, not a finding. **When the program under review is itself a
    sub-program, run (ii) in reverse over its callers** — the index's `callers` lists every row of another book
    that names the ID (`SXJCB178` has 14 calling books, `SXJCB147` 13, `SXJCB141` 12); .xlsx files are zipped, so a
    plain grep finds nothing.
    - **(ii) Sub-program calls.** Caller: a 機能定義書/画面設計書 table `No.|名前|引数名|設定値`, usually under
      「ﾌﾟﾛｸﾞﾗﾑ間の引数定義(大文字/小文字を識別する)」. Callee: its 機能定義書 Ⅳ table `No.|名前|引数名|説明|省略時`
      (`SXJCB121` `[69]`-`[150]`; `SXJCB123` heads it `項目`) and `No.|名前|戻り値|説明` (`[180]`-`[184]`). Identify the callee from
      the 「…(<ID>)を起動」 line above the table (join the two rows above — `PXJCO124` `[167]` 「…(SXJCB141)を」 / `[168]` 「起動する。」), and confirm it by the 引数名 set — the nearest ID can be wrong (`PSJCO201`
      `[298]`/`[300]` name SXJCB141 and SXJCB147 for one table; `PXJCO161` `[776]` mentions SXJCB122 above a SXJCB178 table). Flag:
      - **引数名 differing only in case/underscore — 中** (names are case-sensitive): `PXJCO121` `[96]` `ExclusiveFg_TXJCM003` vs
        `SXJCB121` `[81]` `ExclusiveFgTXJCM003`, and `[125]`… `Meisaikomoku1`-`20` vs `[110]`… `MeisaiKomoku1`-`20` (also PSJCO201,
        PSJCO203); `PXJCO124` `[106]` `MotoEdaBan` vs `SSJCB211` `[70]` `MotoEdaban`; `PXJCO192` `[178]` `HinmokuCD` vs `SXJCB147`
        `[118]` `HinmokuCd`;
      - **an argument the callee does not define, and a callee 省略時=不可 argument the caller omits — 中** — usually one rename
        seen from both sides: `PXJCO161` `[877]`/`[927]`, `PSJCO307` `[348]`, `PSJCO308` `[510]` pass `ShochiKoteiOutputKbn`, while
        `SXJCB141` `[124]` defines `ShochiOutputKbn` (省略時 不可) — `PXJCO124` `[187]` uses the current name; `PSJCO204` `[102]`
        `ShikibetsuSU` vs `SXJCB147` `[113]` `ShikibetsuCardOutputSu`. An omitted argument whose 省略時 is `NULL`/`空白`/a value is
        not reported;
      - **order** (低) and **1-N vs single** (要確認) between the two tables — none found in 2026-10, keep the rule;
      - **an empty 設定値 for an argument whose 省略時 is 不可 — 要確認**. A 設定値 that is only `※n` is not empty: read the
        note for the value's source (`PXJCO161` `[458,23]`/`[461,23]` `※4` → `[516,7]` 「G8)分割領域のみに入力有の場合…」, for
        `SXJCB122` `[90]`/`[93]` 不可 — fine); a list parent row (`KanriNoList` with its children indented below) has no
        value of its own;
      - **return values — 中 when not received, 要確認 when the use is not stated.** Within the step that contains the call (up to
        the next sibling step heading — and including that next step when its heading is a 「戻り値」 one: `PXJCO124` is laid
        out (2)呼出→(3)戻り値, (4)→(5), (6)→(7)戻り値 `[258]`-`[262]`; not a fixed row window), the caller must receive the callee's 戻り値 (ErrList at least) and say what happens: error →
        stop/rollback/display, Url → how it is opened (`[261,21]` 「ﾀﾞｳﾝﾛｰﾄﾞURL 又は URLｽｷｰﾏの文字列を取得する」 with only the error case after it at
        `[262]` is 要確認; also A-⑥ 再印刷 `[363]`/`[364]`), ResultList/KanriNo → where it is used;
      - **a callee definition without 引数名 — 要確認 on the callee** (`SXJCB178` `[47]`-`[62]` is `項目|説明|省略時`, so its 13
        callers' `NyuryokuShubetsu`…`KoteiCd` cannot be verified); also confirm that the callee body still uses each argument.
      Out of scope: `PSJAO702` (【共通】ﾀﾞｳﾝﾛｰﾄﾞ) argument tables (71 of 150) — `file-output-spec-check` owns them. A table
      `…|引数名|…` under a 「戻り値」 heading is the caller's return-value table, not a call (`PXJCO165` `[109]`). Confirmed on
      validation: `PXJCO124` `[106,14]` `MotoEdaBan` vs `SSJCB211` `[70]` `MotoEdaban` — the caller's own return table `[134,14]`
      writes `MotoEdaban`, so it disagrees with itself too.
    - **(i) Screen transitions.** Caller: a Ⅳ event (or 機能処理概要 step) naming another screen's ID with
      `遷移|起動|表示する|開く` and an explicit list — `<受渡ﾊﾟﾗﾒｰﾀ>` / `【受渡ﾊﾟﾗﾒｰﾀ】` / `ﾊﾟﾗﾒｰﾀ⏎・<項目>:<値>`. Callee: its
      `画面設計書(<ID>)` Ⅱ `(1) ﾊﾟﾗﾒｰﾀ定義` (`No.|ﾊﾟﾗﾒｰﾀ|説明／概要`, `※共通` in col 11) from the index. Flag: a passed item missing
      from (1) — 中 (`PXJCO130` `GXJC130A` `[207,20]` / `GXJC130B` `[668,20]` pass `管理No` to `PXJCO161` `画面設計書(GXJC161A)`,
      whose (1) `[63]`-`[72]` has no 管理No); order — 低 (`PXJCO128` `GXJC128A` `[379,23]`/`[444,23]` ﾓｰﾄﾞ→ﾛｯﾄ停止ｷｰ vs `GXJC128B`
      `[109]`/`[110]`); value-source wording outside `ﾛｸﾞｲﾝ情報.`/`ﾊﾟﾗﾒｰﾀ.`/`画面.`/`遷移元画面.`/`遷移元.`/`呼出し元画面.`/`<画面名>.`
      (`加工手順詳細.X`, 91 rows) — 低 (`GXJC134A` `[68]`-`[71]`
      `元画面.会社ｺｰﾄﾞ`…, rows marked `※共通` in col 11), and a callee (1) saying `ﾛｸﾞｲﾝ情報.X` where the caller passes `画面.X` or the reverse —
      要確認. A non-共通 (1) item the caller does not pass is 要確認 only, and silent when the omission matches a mode
      (`GXJC128A` `[378]` 新規 passes only `ﾓｰﾄﾞ:"0"`). Prose transitions (「選択された明細の情報を基に…へ遷移」 — ~160 cells) and
      callee rows reading `遷移元画面.X` (188) are compared by reading, not by script: does the caller's Ⅴ have X, and does the
      prose pass it?
    Partial mechanisation is expected: (ii) is table-to-table and nearly mechanical; (i) is mechanical only for explicit lists
    (16 `<受渡ﾊﾟﾗﾒｰﾀ>` blocks in 2026-10). Cost: other books read per REV — median 1, mean 1.9, 46 of 122 need none. Measured
    2026-10-17: 75 sub-program call tables compared (6 case/spelling, 9 rename/extra/missing-mandatory, 5 empty-mandatory,
    5 return-value candidates of which 2 were handled later), 10 explicit transition lists (2 passed-not-defined, 2 order,
    1 wording).

**Hidden leftovers.** This check also writes the one 【非表示】 要確認 line for every leftover/withdrawn hidden sheet
other than 更新条件表 / 帳票設計書 / ﾌｧｲﾙ仕様書 (画面設計書, ﾁｪｯｸ処理設計書, 機能定義書, 画面ｲﾒｰｼﾞ, scratch copies — the
`agent-guide.md` ownership rule), naming the evidence. Hidden rows holding an old copy of whole sections (`PSJCO404` `GSJC404A` rows
711-1413 = an old Ⅲ (1)-(6) with placeholder `TSJCDXXX`, rows 1625-1662 = an old Ⅳ) and any other hidden
body text no check's subject covers are reported **here**, once per sheet, 【非表示】 要確認, every range in
the one line: 「非表示の行 711-1413, 1625-1662 に旧版の Ⅲ/Ⅳ が残っている。不要なら削除、必要なら再表示」;
no other check — of this skill or another — reports what is inside them. A **live** reference whose only
target is hidden — Ⅴ `[1693,35]` 「Ⅲ．画面表示仕様>(1).検索時数量設定表示有無」, a name found only in hidden row
`[733,6]` (the live (1) `[109,8]` is 「検索時数量設定」) — is a dangling reference, reported by whichever check
resolves that kind of reference (here, as part of the same line when no other does). Rows holding only
change-history memos are skipped (`agent-guide.md`, "Hidden sheets, rows and columns").

**Shapes that are not defects (from the 2026-10 validation runs).**
- An area can have **several Ⅵ control tables** (one per mode or per tab); merge them before
  deciding an item has no control row.
- A Ⅵ area written as **one delegation line** (`G5)実績表項目領域 ※XJC_ｼｽﾃﾑ共通設計書(実績表項目設定)参照`,
  `PXJCO161` `GXJC161B` `[3611,3]`, `[4026,3]`) has its control spec elsewhere — not "area has no Ⅵ block".
- A **grouped name** covers a series: `直行1-4` in Ⅵ (or Ⅳ) stands for `直行1`…`直行4` in Ⅴ; so does a
  slash-joined unit series (`不良数(ｼｰﾄ数/作業基板数/基板数/個数)`, `GXJC161B` `[1852]`-`[1855]`) and a Ⅳ name
  written without its `(1-N)`.
- Ⅰ．画面ﾚｲｱｳﾄ is a floating picture (usually EMF), so its empty cells in the dump are not a
  missing layout. For check 5's three-way match, get the picture via `agent-guide.md` "Seeing the
  actual screen layout" (unzip `xl/media/*.emf`, convert to PNG, Read it — no Excel); this completed
  the Ⅰ/Ⅴ/Ⅵ comparison on `PXJCO125`. Report only items clearly drawn or clearly absent, as a
  visual comparison. When the workbook has several screen sheets, map picture → sheet from the zip,
  not COM: `xl/workbook.xml` (sheet name → `r:id`) → `xl/_rels/workbook.xml.rels` (→ `worksheets/sheetN.xml`)
  → `xl/worksheets/_rels/sheetN.xml.rels` (→ `drawings/drawingN.xml`) → `xl/drawings/_rels/drawingN.xml.rels`
  (→ `media/imageK.emf`). Worked on `PSJCO403` and `PXJCO161`.
- A digest row that looks deleted may have been **moved** (struck at the old place, live elsewhere);
  search the live dump for its text before reporting a removal-driven asymmetry.

## Reporting

The output goes to the designer who owns the doc, not into your own working notes — report only
what they'd actually need to act on.

For each real finding: state the defect in one or two sentences, cite the exact sheet name and cell
coordinates (and the master file it was checked against, e.g. 82.画面項目辞書, 04.メッセージ管理,
06-09_レスポンス一覧) so it's actionable, and give a concrete suggested fix when one is
obvious (e.g. "change X to Y"). Group findings under a checklist-item
heading (e.g. "3-4: ...") only when several findings share one; don't force every finding into the
checklist's numbering if it doesn't fit cleanly.

Order by impact, most important first — a broken cross-reference, a missing table, or a functional
gap belongs well before a stray full-width/half-width punctuation difference or a naming nit. When a
finding depends on a judgment call rather than a clear rule (a naming convention that might be
intentional, a cross-WG reference you couldn't fully verify, a pattern that could be deliberate
delegation to a common design doc rather than an omission), say so explicitly and mark it as needing
the designer's confirmation — don't present it with the same confidence as a confirmed defect.

Omit entirely, unless the user explicitly asked for a full pass-by-pass audit:
- "確認済み・問題なし" lines for checks that simply passed
- tallies of how many items/tables/IDs were checked
- narration of your own process (which files you dumped, how you sampled, which scratch path you
  used)

The one exception worth keeping brief: if a whole section couldn't be verified at all (a master file
was missing, a section was empty in this program's doc), a one-line note is fine — that's a real gap
in review coverage the designer should know about, not a process detail.

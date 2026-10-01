---
name: design-doc-internal-consistency
description: Review one program's design-doc set (機能定義書/画面設計書/ﾁｪｯｸ処理設計書/更新条件表) for internal cross-reference consistency — e.g. every screen event is reflected in the processing overview, response definitions are registered, every screen-item ID / message ID is registered in its master list, screen layout/item-definition/control-spec agree (including item ORDER matching between Ⅳ．画面項目ｲﾍﾞﾝﾄ詳細 and Ⅴ．画面項目定義, not just which items are present), exclusive-control is present when a program updates a table, process/check order vs button order, and 区分名称 references resolving in 09.区分名称_step2.xlsx. The self-check writing rules (Ⅴ per-item rules, 項番) live in `design-doc-writing-rules`. Grounded in this project's own official review checklist. Does NOT cover the Ⅲ．入出力定義 (I/O table) completeness check — that's the more involved `design-doc-io-table-check` skill, split out separately. Use when the user asks to レビュー/REV a program's 機能定義書 or 画面設計書, or asks whether a design doc is "internally consistent" / "漏れがないか", or whether screen-item ordering matches across sections. For a full REV without named checks the entry point is `rev-program-review`; run this standalone only when selected there or asked for by name.
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
   checklist `[19,4]`). Collect the 画面設計書 "Ⅳ．画面項目ｲﾍﾞﾝﾄ詳細" rows whose ｲﾍﾞﾝﾄ内容 (col 13) is
   `起動時` or `ﾎﾞﾀﾝ押下`, in every screen area, and every step of 機能定義書 "Ⅳ．機能処理概要". Each such
   event should have a traceable step (wording need not match). Flag any with none. `ﾛｽﾄﾌｫｰｶｽ`,
   `ﾀﾞﾌﾞﾙｸﾘｯｸ` and similar events are outside 3-4 — flagging them gave three false positives on
   `PXJCO125` `画面設計書(GXJC125A)` `[415]`/`[417]`/`[418]`. A zoom button (`作業場ｽﾞｰﾑ` `[416]`) is a
   ﾎﾞﾀﾝ押下 event, but process overviews routinely leave zooms out: report a missing one as 低/要確認.

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
   - **A prefix belonging to a WG other than the one you're reviewing and other than 共通** (e.g. an
     `SJA-xxxxx` id inside a 工程管理 program): don't assume it's simply missing — note it
     separately as "cross-WG reference, couldn't verify without checking that other WG's own
     message file" rather than a confirmed gap.
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

5. **4-9/4-13 — Screen layout ↔ item-definition ↔ control-spec three-way match.**
   Collect item names from "Ⅰ．画面ﾚｲｱｳﾄ" (the visual mock), "Ⅴ．画面項目定義" (the item table), and
   "Ⅵ．画面項目制御・出力仕様" (the control matrix). All three should list the same set of
   non-hidden items. Flag any item present in one but missing from another (hidden items are
   expected to be absent from Ⅵ per the checklist note — don't flag those).

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
   to the `yyyy/MM/dd` convention) as a casing defect to correct.
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
   hidden items from this comparison the same way check 5 does (they're not always listed in every
   section). When you can't tell which of the two orderings is authoritative (neither section
   carries an explicit "this is the source of truth" marker), report the finding as a plain "order
   mismatch" between the two locations without asserting which side is wrong, and let the designer
   decide which one to resequence.

7. **5-4 — Exclusive-control checks present when updating.**
   If any 更新条件表 sheet exists for this program (i.e. it writes to a table), confirm
   ﾁｪｯｸ処理設計書 mentions a 排他ﾁｪｯｸ, and confirm "Ⅴ．画面項目定義" has a hidden 排他ﾌﾗｸﾞ item. Flag
   if a program updates a table but has neither.

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
     `削除`; (3) a tie of equal length is left unmatched. Never match a heading that names a link or a
     frame/row operation inside an area — `…ﾘﾝｸ押下処理`, `承認ｸﾞﾙｰﾌﾟ枠削除処理`, `通知先行削除処理` — those are
     not 共通 buttons even when a 共通 button's name (`承認`, `削除`) appears in them. Items that match no 共通 button (links, zoom, other-area
     buttons) are ignored, not reported. Then flag any pair of matched items whose relative order
     differs from the reference.
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
    - **Parsing the reference**: take the text after `区分名称.` up to `参照`/whitespace; drop an
      unbalanced trailing `)` (it closes an outer `区分名称(ｼｽﾃﾑ共通設計書.区分名称.製造ｵｰﾀﾞｰ状態)`); keep a
      `(qualifier)` as part of the name; resolve `group.value` (value must be one of the group's names) and
      `group.qualifier` → `group(qualifier)` before calling anything missing.
    - Measured: 525 references, ~441 resolve once the index handles both title shapes. Report:
      - **placeholder `区分名称.XXXX`** — 中 (≈30, `PSJCO403`/`404`/`405`, plus `検索区分(XXXX)` in
        `PSJCOB01`); the 区分 was never decided;
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
        (`PXJCO125` `画面設計書(GXJC125A)` `[478,43]`), whose group `区分名称.ﾛｯﾄ構成区分` is in `[478,25]`;
      - **qualifier missing or different** — 要確認 (11): a bare `検索区分` when the master has
        `検索区分(ﾛｯﾄﾄﾚｰｽ)`/`(不良ﾛｯﾄ統合)`/`(社内加工用)` — say which one is meant; `検索区分(社内加工)` vs the
        master's `(社内加工用)` is a near-miss to correct.
    - **Inline value lists** (`'0':表示しない、'1':表示する`, or an 初期値 like `11:新規発行` / `0:通常`, beside an item named after a group): compare
      code and name with the master when present. Only 7 rows in the corpus carry one and all matched, so
      expect silence.
    - The same scan surfaces `ｼｽﾃﾑ共設計書`(sic — `通` missing) in many references; that is a typo for
      `design-doc-typo-check`, mention it only if that check is not in the run.

**Shapes that are not defects (from the 2026-10 validation runs).**
- An area can have **several Ⅵ control tables** (one per mode or per tab); merge them before
  deciding an item has no control row.
- A **grouped name** covers a series: `直行1-4` in Ⅵ (or Ⅳ) stands for `直行1`…`直行4` in Ⅴ.
- Ⅰ．画面ﾚｲｱｳﾄ is a floating picture (usually EMF), so its empty cells in the dump are not a
  missing layout. For check 5's three-way match, get the picture via `agent-guide.md` "Seeing the
  actual screen layout" (unzip `xl/media/*.emf`, convert to PNG, Read it — no Excel); this completed
  the Ⅰ/Ⅴ/Ⅵ comparison on `PXJCO125`. Report only items clearly drawn or clearly absent, as a
  visual comparison.
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

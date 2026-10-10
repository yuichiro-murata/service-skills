---
name: design-doc-typo-check
description: Proofread the free-text Japanese prose inside a program's design-doc workbook — feature/processing narratives, event descriptions, error-message text, footnotes, revision-history notes — for actual language mistakes (誤字/脱字/衍字/助詞の誤り/変換ミス), as distinct from its sibling skills' structural checks: cross-reference completeness (design-doc-internal-consistency), I/O table completeness (design-doc-io-table-check), DB column existence (xlsx-db-column-check), ID-numbering rules (naming-standard-compliance), and font/merge irregularities (design-doc-formatting-consistency). Use when the user asks to check 誤字脱字 or wants a language-proofreading pass on a design doc. For a full REV without named checks the entry point is `rev-program-review`; run this standalone only when selected there or asked for by name.
---

# design-doc-typo-check

Proofreads one program's design-doc workbook for genuine Japanese-language errors in its free-text
prose — wrong kanji, a missing character that breaks the sentence, a stray duplicated character, a
wrong particle that flips the meaning, a garbled leftover fragment from a copy-paste edit. This is
fundamentally different from every sibling REV skill: those are mechanical diffing/lookup tasks;
this one requires actually reading each sentence and understanding what it says. Do not try to
shortcut it with keyword search or pattern matching — read the prose.

This skill does **not** check cross-document terminology consistency, ID format, or font/merge
signals — see the frontmatter `description` above for which sibling skill covers each. A formatting
anomaly `design-doc-formatting-consistency` flags is often a good hint of *where* a copy-paste-era
typo might also be hiding, though, so it's worth cross-referencing that skill's findings if run in
the same pass.

## Environment

**Read `_shared/agent-guide.md` first** — scope selection, environment, the live dump, reporting
conventions. This check reads only the target workbook's dump: everything in the `.txt` is live
prose, and deleted-in-spirit text was never handed to you to proofread.

This matters more to proofreading than to any sibling check, because the artefact it removes looks
exactly like this skill's own quarry. A partially-struck cell arrives already reduced to its live
text, so a cell that would have read as a **garbled leftover fragment** in a raw dump now reads
correctly. Judge what you see; do not reconstruct the original from `_DELETED_DIGEST.txt` and
proofread that. Confirmed false findings from raw dumps: `"GSXJC205A"` reported as a misspelling of
`GSJC205A` (only the `X` was struck), and `注意事項備考` reported as a nonexistent term (only
`注意事項` was live).


## Scope: which cells are "prose" here

Structured tokens — IDs, table/column names, `<alias>.<column>` expressions, numeric codes, half-width
katakana terms that are this project's standard house style (`ﾃﾞｰﾀ`, `ﾛｯﾄ`, `ｵｰﾀﾞｰ`, `ｺｰﾄﾞ`, etc.) —
are NOT prose and are NOT this skill's concern; those belong to the other REV skills or are simply
normal project vocabulary. One exception on every sheet type: a structured token is still in scope for
an **obvious** misspelling (`ﾛｸﾞｲﾝ画面.` with a stray period, `作場名` for `作業場名`, `祖No` for
`層No`) — skip only style and abbreviation questions there. Otherwise only free-running Japanese
sentences and phrases are in scope. In a typical program workbook, that means:

- **表紙**: Ⅲ．改訂履歴 の 改訂内容 column (what each revision actually changed).
- **機能定義書**: Ⅰ．機能概要 (prose bullets), Ⅲ．入出力定義 の 用途 column (short prose),
  Ⅳ．機能処理概要 (the main narrative — numbered steps, ※ footnotes, sub-program 設定値/戻り値 notes,
  pattern captions, and any prose woven into 検索条件/取得内容 descriptions), Ⅴ．ﾌﾟﾛｸﾞﾗﾑ構成 (ｴﾗｰ時処理),
  Ⅵ．前提条件, Ⅷ．その他特記事項 (if present). (Ⅱ is `Ⅱ．I/O関連図`, a picture — nothing to read.)
- **画面設計書**: all prose in Ⅱ．画面補足説明・表示ﾊﾟﾀｰﾝ概要 through Ⅵ — Ⅲ．画面表示仕様's numbered
  steps as well as its ※ footnotes, Ⅳ．画面項目ｲﾍﾞﾝﾄ詳細 (event-behavior descriptions — often long
  paragraphs), every prose column of Ⅴ．画面項目定義 — 説明 (col 43; there is no 備考 column), 初期値
  (col 35/36, often multi-line case text: `PXJCO161` `GXJC161D` `[291,35]`) and 画面項目名 (col 5, obvious
  misspellings only), located by header label since they shift by one — and the ※-notes under
  each Ⅵ．画面項目制御・出力仕様 table (cols 3/5, e.g. `PXJCO125` `GXJC125A` `[582,5]`; Ⅵ's own columns
  are only No./画面項目/one per event).
- **ﾁｪｯｸ処理設計書**: the `ﾁｪｯｸ内容・経緯` column (col 14, header `[8,14]` — most of the sheet's
  prose; `PXJCO125` `[35,14]`/`[37,14]` findings), the `ﾁｪｯｸ詳細` (col 59) and `ﾁｪｯｸ補足` (col 86)
  columns, error-message body text — **checked against the registered text** (step 2, last bullet).
- **ﾌｧｲﾙ出力仕様書 / ﾌｧｲﾙ入出力仕様書 (titled ﾌｧｲﾙ入力仕様書)** (if present): the Ⅰ．ﾌｧｲﾙ出力/入力条件
  narrative (FSJC018 `[8,3]`), 出力条件 / 編集内容 / 備考 prose, and ※-notes.
- **更新条件表**: find the cells by role, not coordinate — a sheet repeats its block once per update
  pattern (`PXJCO161` `TSJCA006`: 8 blocks, labels `[9,2]`, `[109,2]`, `[209,2]` …). In **every** block: the
  更新概要 cell (value right of the `更新概要` label — a ①②③ source list, where wrong screen IDs hide —
  `PXJCO125` `TSJCD014` `①画面(GSJC201B)`), the **更新条件** narrative (right of the `更新条件` label), ※-notes
  in the value half of the `INSERT`/`UPDATE` column (col 18 — `TSJCA057` `[28,18]`), ※-numbered footnotes
  (written in full sentences, unlike the 項目名/取得内容 columns which are mostly structured tokens), and
  the right-hand 参照ｴﾝﾃｨﾃｨ/notes panel (col 54 onwards up to the revision margin — `TSJCA057` `[18,54]`
  `ﾛｯﾄ情報の更新後に取得する。`, `[74,59]`, entity captions in col 68).
- **帳票設計書** (if present): 処理の流れ (numbered narrative steps), 備考 column.

- **Supporting sheets the spec relies on**: a sheet that is not a doc type in 表紙 Ⅱ but is cited from
  the spec (`【<sheet>】参照`, `「<sheet>」ｼｰﾄ参照`) or listed under 表紙 Ⅱ's その他設計書 — `PXJCO161`
  `計算ﾎﾞﾀﾝ押下時の計算方法` (cited by 機能定義書 `[999,5]` and `GXJC161B` `[1856,24]`; it held a real typo),
  `PSJCO403` `自動ｾｯﾄの動き` (cited by `GSJC403A` `[1240,30]`). Grep the spec sheets for each extra sheet name.
- **Designer memos right of the form** (画面設計書 cols ≥ 55, e.g. `GSJC403A` `[47,55]`): proofread them
  and use them as evidence of intent.

**Out of scope**: designer scratch sheets that are neither a doc type in 表紙 Ⅱ nor cited from the spec
(PSJCO403's `数字を調整の流れ`, `検索SQL検証`, `資料`, the `…old` copies — the dump includes them; that
they are left visible is `design-doc-formatting-consistency`'s call, not a typo); SQL text wherever it
sits, including its `--` comments (GSJC403A `[281,54]` in Ⅲ, `[219,55]` — col 54 is otherwise the
画面項目ID column of Ⅳ/Ⅴ, not an SQL column); the revision margin (cols 105-112, `2026/02/12 平島 追加…`),
which is change history.

## Procedure

1. Use the dump `rev-program-review` handed you; standalone, run
   `python _shared/scripts/live_dump.py <workbook> <out_dir>` (it skips 詳細設計書). A large dump
   (`PXJCO161` `GXJC161B` is 640 KB) cannot be read whole: load it with `_shared/scripts/dump_cells.py`,
   keep the in-scope cells that contain Japanese text, de-duplicate identical values (keep every
   coordinate for the report), and read that list.
2. For each section listed above, read every cell's full text — sentence by sentence, not a keyword
   scan. Look for:
   - **誤字**: a wrong kanji/character that doesn't fit the intended word (a likely IME conversion
     slip, e.g. a homophone substituted for the correct word).
   - **脱字**: a missing character or word that makes the sentence ungrammatical or leaves an action
     without its object/verb.
   - **衍字**: an accidentally duplicated character or word (e.g. "確認するする").
   - **助詞の誤り**: a wrong particle that changes or breaks the meaning (e.g. "を" where "が" was
     needed, an action that reads backwards from what's clearly intended).
   - **Unmatched brackets/parentheses/quotes** within one cell — except notation tokens (`G1)`,
     `1)`, `(1)①`, `②.2`) and a bracketed expression that deliberately spans cells or rows (a `(` in one
     検索条件 row closed by `)` in a later one, `GXJC161A` `[369,5]`/`[370,48]`).
   - **Garbled leftover fragments**: a word or clause that doesn't belong to the sentence it's sitting
     in — often the tell-tale sign of a copy-paste edit where the old text wasn't fully replaced.
   - **Error/confirm message text vs the message ledger.** Before reporting punctuation, wording or a
     typo in a message quoted in ﾁｪｯｸ処理設計書 / Ⅳ, look up its ID in `01_Doc\04_共通設計\04.ﾒｯｾｰｼﾞ管理_<WG>.xlsx`
     (`XJZ`/`SJZ` → `_共通`; read live with `live_dump.py --sheets "^XJZ\("` etc.; the text is in col 13).
     If the design doc's text equals the registered text, report nothing about its punctuation or
     style — the screen shows the ledger's text (`XJZ-000034` `対象ﾃﾞｰﾀを選択してください` has no `。` in
     the ledger either; PSJCO204 9/30 6-4/6-5 and PSJCO304 9/14 6-11 were rejected for this). Report only
     a difference from the ledger (as a mismatch, the ledger being the reference), or an obvious
     misspelling present in both (once, noting it is the ledger's text). A message with no ID, or whose
     ID is in no ledger sheet (confirm-message IDs such as `XJZ7006` live in 82.画面項目辞書, not the
     ledger), is proofread normally; whether the ID is registered is `design-doc-internal-consistency`'s.
3. Deprecated rows were already dropped at dump time — do not scan for them, and do not treat a
   sentence that reads oddly as evidence of a leftover you need to go verify in Excel.
4. Cross-check every candidate against context before reporting it as a real error:
   - **Is it house usage? Measure before reporting.** For any wording/notation candidate (not a plain
     wrong kanji), grep the self-check workbook `01_Doc\08_機能定義書\11_工程管理\XXXXX000_ｾﾙﾌﾁｪｯｸ用.xlsx`
     and the PHASE corpus for both the written form and your proposed correction (live text — e.g. a
     python loop over `live_dump.classify` with openpyxl `read_only=True, rich_text=True`). "The corpus"
     = the workbooks directly in `PHASE1`-`PHASE3`, excluding the target. If the corpus uses the written
     form at least as often as the correction, it is house usage: do not report it — also when the same
     workbook mixes both forms (PSJCO304 has `閉開` and `開閉`). The self-check sample decides only when it
     uses the written form alone (it has both `閉開` GXJA802A `[296,24]` and `開閉` GXJA802B `[102,24]`). For
     a candidate with no greppable correction (punctuation, a quoting style such as
     `"1"：実績出力するを設定する`), count the pattern itself: if it recurs across several workbooks
     (`出力するを設定する` 20 in 7, `をﾁｪｯｸ時` 53 in 32, `引き渡しされる` 187 in 80), it is house style. Designers rejected these repeatedly (PSJCO204
     received `閉開→開閉` four times, PSJCO304 twice, all answered 対応不要). Confirmed house forms
     (2026-10-10, 122 PHASE1-3 workbooks): **`閉開`** (ｱｺｰﾃﾞｨｵﾝ閉開; 92 occurrences in 59 workbooks vs `開閉`
     12 in 11; the self-check sample uses `閉開`) and **`文字列(半数)`** (18 in 10 workbooks; `文字列(半数字)` /
     `文字列(半角数字)` 0 — designers cite the self-check sheet for it).
   - **Terse noun phrases are not 脱字.** `明細ﾍｯﾀﾞｰをﾁｪｯｸ時`, `検索・実行時`, `〜を<名詞>時` are the
     project's compressed style for event/condition labels (PSJCO204 9/30 6-2, 6-6, both rejected).
     Report a 脱字 only when the sentence becomes unreadable or loses its object/verb.
   - **表紙 Ⅲ．改訂履歴 is a historical record.** Proofread its 改訂内容 for obvious misspellings only;
     never compare it with today's 項番, names or IDs (a revision note naming a since-renamed item is
     correct as history — PSJCO304 9/28 6-3, rejected).
   - Does a near-identical sentence exist elsewhere in the same doc (a sibling row in a repeating
     list, a parallel branch for a different condition)? Diffing against that twin is often the
     fastest way to confirm a real error versus a deliberate variation.
   - Could ambiguous phrasing be intentional shorthand common in this project's docs, rather than a
     mistake? If genuinely unsure, report it as a lower-confidence note rather than a confirmed error
     (see Reporting).
5. Hand off what is not a language error (one line each, not reported here):
   - a stale but real ID (`GXJC163B` in `ﾁｪｯｸ処理設計書(GXJC161B)`, `機能定義書(PSJCO203)`, `VDMBM04_31`) →
     `naming-standard-compliance` / `design-doc-internal-consistency`;
   - a leftover name after a rename (`処理取消` for a renamed button/event) → `design-doc-internal-consistency`;
   - a swapped pair of terms (`依頼先`/`依頼元`) → `design-doc-internal-consistency`, unless it is purely a
     misspelling of the word itself;
   - the revision margin (cols 105-112) → out of scope.

## Reporting

The output goes to the designer who owns the doc, not into your own working notes — report only
what they'd actually need to act on. Write in Japanese.

For each real finding: cite the exact sheet name and cell coordinate, quote the erroneous text
verbatim, state plainly what's wrong (誤字/脱字/衍字/助詞の誤り and which), and give the corrected
text when it's obvious. Order by impact: an error that changes or obscures the meaning of a processing
rule or condition belongs before a cosmetic misspelling that doesn't affect understanding. Distinguish
confirmed errors from "this reads oddly but might be intentional shorthand" judgment calls, and mark
the latter as needing the designer's confirmation rather than asserting it as a defect.

Omit entirely: a tally of how many cells/sections were proofread, "確認済み・問題なし" notes for
clean sections, and narration of your own process (which files you dumped, how you sampled). The one
exception worth a one-line mention: a whole prose section that couldn't be read at all (e.g. it was
entirely struck-through, or a sheet's used range exceeded the dump cap) — that's a real coverage gap
worth flagging, not a process detail.

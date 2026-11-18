# Fixtures

Every `.docx` here is opened, projected, exported and compared part by part by
`test/roundtrip.test.mts` (ED-002 section 3.7), and its opaque report is pinned in
`diagnostics.expected.json` (`test/diagnostics.test.mts`; regenerate with `UPDATE_DIAGNOSTICS=1`
and say why in the commit message). The edit tests name the fixtures they use.

Copied, not referenced across repositories, so the package's tests run from npm.

Embedded fonts: `fonts-embedded.docx`, `fonts-embedded-subset.docx` and `fonts-embedded-subset-m365.docx`
embed Noto Sans, under the SIL Open Font License 1.1; its copyright notices and text are in `OFL.txt`
here. Where the text extraction suite is present (below; the public repository leaves it out),
`text-extraction/suite/phase0/extraction-probe-inspect.docx` embeds fonts cut from Liberation Sans, also under the
OFL, whose notice docx4j keeps beside them and the copy brings (`text-extraction/suite/support/inspect-fonts/OFL.txt`).

| Fixture | Source | Why it is here |
|---|---|---|
| `HelloWordOnline.docx` | `docx4j-core-ts/test/fixtures` | Word Online output; the main part is `word/document22.xml`, so nothing may assume the part's name |
| `loadAndSave.docx` | same | the engine's own round-trip fixture; its author (properties, comment, people part with its account), its sensitivity label's ids and its chart workbook's author replaced 2026-11-19 |
| `hyperlink_dupe.docx` | same | a TOC: hyperlinks around runs, a tab and a PAGEREF field (holders, opaque field runs, an empty run); its author and company replaced 2026-11-19 |
| `header-no-rels.docx` | same | a header part with no relationships part |
| `invoice.docx`, `invoice2013.docx` | same | OpenDoPE content controls, block and run level, nested; row-level controls around table rows (nested in `invoice2013.docx`, whose second table has no `w:tblGrid`) |
| `tables.docx`, `MergeWithNext.docx` | `docx4j/docx4j-samples-docx4j/sample-docs` | tables (E2.a): vertical and horizontal merges, header rows, a 13-column grid, table styles; `MergeWithNext.docx` a plain two-by-two; `tables.docx`'s author, company and its picture's link path (and alt text) replaced 2026-11-19 |
| `trailing-table.docx` | Jason, 2026-10-08 (`test_bind (1).docx`) | a body that ends with a table, its last row in a row control, and no paragraph after it (Word adds one on open): the gap cursor, ED-005 section 12.73 |
| `tracked-changes.docx` | `docx4j-core-ts/test/fixtures` | `w:ins` shown as text, `w:del` hidden but carried (decision 2); a table with a vertical merge |
| `comments-two.docx` | same | comment ranges as marks, the reference runs opaque |
| `comments/comments-for-word.docx` | written by the editor in a room (`build/comments-for-word.mjs`, 2026-09-29) from `tracked-changes.docx` | E4.e's comments as the editor writes them: made, replied to, resolved and deleted, with small ids, every comment part and both comment styles made by the editor (ED-005 section 12.22) |
| `comments/comments-for-word-word15.docx`, `comments/comments-for-word-word2010.docx` | that file opened and saved by Jason in Word 15 and in Word 2010 (2026-09-29) | what each Word keeps of the editor's comments: every comment, the ids renumbered in document order; Word 15 the threads and paragraph ids; Word 2010 no paragraph ids, the side parts kept as they were (ED-005 section 12.23) |
| `comments/comments-for-word-word15-compat.docx`, `comments/comments-for-word-word15-upgraded.docx` | the editor's file opened again in Word 15 by Jason and saved untouched, in compatibility mode and converted (2026-09-29) | the resolved state kept in compatibility mode and cleared by the conversion to Word 15's format (ED-005 section 12.23 item 3) |
| `comments/comments-for-word-word365-compat.docx`, `comments/comments-for-word-word365-upgraded.docx`, `comments/comments-for-word-word15-compat-word365.docx` | the editor's file, and Word 15's compatibility-mode save of it as a control, saved by Jason in current Word (Microsoft 365, 2026-09-29) | current Word showing the resolved state and saving it in neither form, whoever serialised it (ED-005 section 12.23 item 7) |
| `comments-span.docx` | derived from `comments-two.docx` (2026-09-26, by a script rewriting `document.xml`) | comment 0's range over three paragraphs: a tree run through its middle paragraph in a room (ED-004 section 3.10) |
| `fonts-hebrew-no-cs.docx`, `fonts-cjk-linebox.docx`, `fonts-symbol-and-emoji.docx` | `docx4j-core-ts/test/fixtures/parity` | Hebrew, Japanese and emoji text (ED-001 section 3.7) |
| `numbering-stories.docx`, `styles-linerule.docx` | same | lists (E2) and line spacing: the paragraphs must round-trip now |
| `headers.docx`, `images.docx` | `docx4j/docx4j-samples-docx4j/sample-docs` | six header parts over five sections (the stories test); inline images (delete test) |
| `headers-footers.docx` | `docx4j/docs` | sections without header or footer parts |
| `numbering-clauses.docx` | a generated extract of an association's rules Jason supplied (2026-10-14), its four empty relationships parts removed; its text, footer and properties replaced 2026-11-19 by neutral text of the same lengths (the tests cut it at positions) | clause levels linked to styles with a level-1 item before any level-0 one (ED-005 section 12.97); a footer with a right tab stop at the text's edge and a complex PAGE field with no cached result (section 12.98) |
| `sections-multiple.docx`, `odd-even-first-page.docx` | `docx4j/bin-testOutput/headers_footers` | several sections with `w:sectPr` on paragraphs, "same as previous"; first, even and default headers under `w:titlePg` and `w:evenAndOddHeaders`; the final section edit test |
| `footnote.docx` | `Plutext-Enterprise-Java-8/testing` | footnote references and their notes, one of them an orphan |
| `mergefield.docx` | `docx4j/docx4j-core-tests/src/test/resources` | MERGEFIELD field runs |
| `drawing-anchor.docx` | same | an anchored drawing, opaque |
| `bookmark-spans-paragraphs.docx` | same | a bookmark whose range spans paragraphs |
| `w15-repeating-section.docx` | `docx4j/docs/developer` | w15 repeating section content controls beside OpenDoPE ones |
| `header-image.docx` | `docx4j/docx4j-core-tests/src/test/resources/AlteredParts/header-simple-plus-image.docx` | a header with an image through the header part's own relationships (E2.d, ED-003 section 6.6) |
| `notes-numbering.docx` | built by `build/notes-numbering.mjs` with the engine | footnotes and endnotes over two sections with `w:footnotePr` and `w:endnotePr` (format, start, restart each section), `w:pgNumType` with a start and a format, one page-break marker (E2.d: the note numbering and page number display, the endnote edit tests) |
| `print-pages.docx` | built by `build/print-pages.mjs` with the engine | four sections (A4, A4 landscape after a next-page break, the same landscape page after a continuous one, US Letter with a gutter; mirror margins in the settings part); a page break inside a paragraph, a direct page break before, keep with next on and off, widow control off, a footnote and an endnote, a tracked insertion and deletion, a content control showing its placeholder, a table of two header rows, forty rows and a row that cannot split (ED-006: printing) |
| `nested-table.docx` | built by `build/nested-table.mjs` with the engine | a table inside a cell of a table (the nested-table suite, 2026-09-25) |
| `revisions/revisions-base.docx` | built by `build/revisions-base.mjs` with the engine (2026-09-28) | the document Jason revised in Word 2010 and Word 15 on the reference machine, as two authors with tracking on, to make E4.c's fixtures (ED-005 section 13; `revisions/README.md` says what each section yields); it holds no revision |
| `revisions/revisions-word2010.docx`, `revisions/revisions-word15.docx` | made by Jason in Word 2010 and in Word 15 from `revisions-base.docx`, following the steps it carries (2026-09-28) | E4.c's Word-made revisions (ED-005 sections 5.2 and 13): text inserted and deleted by two authors, an insertion deleted by another author, a move with its range markers and moved paragraph marks, run and paragraph formatting changes, a numbering change as `w:pPrChange`, paragraph marks inserted and deleted, rows inserted and deleted with their cells, a deleted tab and an inserted break and note reference, table, cell, grid and section property changes, a revision in the footnotes part; comments by two authors, and in Word 15 a reply marked done (`commentsExtended.xml`, `people.xml`). `test/revisions.test.mts` reads them |
| `revisions/check18/`, `revisions/check20/`, `revisions/check22/`, `revisions/check24/`, `revisions/check25/`, `revisions/check26/` | Word 16.0.20326.20158 saves by Jason, copied from docx4j-core-ts's `test/fixtures/revisions/` (checks 18, 20 and 22 of its `test/README.md`, 2026-09-28) | what Word did when accepting and rejecting: check 18 a move from either half and all at once (from `revisions-word15.docx`), check 20 edited and broken moves (each case's input and its `-accept` and `-reject` saves), check 22 a deleted paragraph, a lone deleted mark and a deletion across a mark; check 24 formatting changed back with tracking on, by its author and by another (core-ts 03b4d9f; `test/track.test.mts`); check 26 the same for a paragraph's alignment and indent (core-ts bce8206; `test/track.test.mts`); check 25's `25a-bullets.docx` and `25b-bullets-reject.docx` (Word 16, core-ts aeef42e) a list turned to bullets and rejected, and `25e-bullet-demote.docx` and `25f-bullet-demote-reject.docx` (core-ts d68e803) an item bulleted and moved a level, rejected. `test/review.test.mts` holds steps C4 and C5 to them (ED-005 sections 12.6 and 12.7). Below `allFixtures`' reach, so the per-fixture suites and the pinned diagnostics do not read them |
| `revisions/check30/`, `revisions/check31/`, `revisions/check32/`, `revisions/check33/`, `revisions/check34/`, `revisions/check35/` | Word 16.0.20326.20158 and 16.0.20430.20118, run by Jason through core-ts's Office JS probe, copied from docx4j-core-ts's `test/fixtures/revisions/` (its commits 426e7b5, 2414220 and 8cfdef9) | `result.json` per check: the markup put in, what Office JS's `body.getTrackedChanges()` listed and the markup Word left after each action. Check 30 tracked insertions holding run holders (`test/check30.test.mts`, ED-005 section 12.40); check 31 inserted rows (`test/check31.test.mts`, section 12.45); check 32 a pending insertion or deletion holding a comment's word resolved, and typing into a link with tracking on (`typed.json` and Word's save `check32-typed.docx`; `test/check32.test.mts`, section 12.49); check 33 Word's saves of a whole heading paragraph deleted with tracking on and then accepted (`33-input.docx`, `33a-deleted.docx`, `33b-accepted.docx`; `test/check33.test.mts`, section 12.51). Check 34 (core-ts 1074cfe, Word 16.0.20430.20118) Word's XML Mapping Pane inserting a bound plain text control, one case per paragraph of the engine-made `34-input.docx`: `34a-mapped.docx` tracking off, `34b-tracked.docx` case 1j with tracking on, `34b-tracked-reject-all.docx` Word's Reject All over it, `34d-mapped-existing.docx` an existing control mapped, `probe.json` the Office JS probe (`test/check34.test.mts`, sections 12.56 and 12.57; no `result.json`). Check 35 (core-ts 158a111 and a672427, part 1 in current Word) rich text controls mapped in Word's XML Mapping Pane: `35-input.docx` made by core-ts's `scripts/make-check35-input.mjs`, `35a-word15.docx` Word's save, `35c-word2010.docx` and `35e-word2010.docx` Word 2010's saves of it and of the editor's `35d-editor.docx` (made by `fixtures/build/check35-for-word.mjs`; core-ts b1b96a3) (`test/check35.test.mts`, section 12.58). Below `allFixtures`' reach |
| `review/` | E4.g's Word session (ED-005 section 9.2, step G1, 2026-09-30): `review/README.md` says what Jason does in Word and which saves come back here. The files for the session are not kept here but built by `npm run word-session` into `packages/editor-mcp/test-results/word-session/`: the editor's file made in a room by two people and the agent (`packages/editor-mcp/scripts/review-for-word.mjs`), the measurements file (`build/measure-base.mjs`) and the id bisection's twelve files (`build/ids-for-word.mjs`); `test/word-session.test.mts` and the companion's `test/word-files.test.mts` hold the builders. The `review-for-word.docx` Jason opens comes back with his saves, since each build has other dates and ids |
| `paste/paste-samples.docx` | built by `build/paste-samples.mjs` with the engine | the document Jason opens in Word to capture the paste corpus's clipboard HTML (ED-003 section 8.3; `paste/README.md` says how); not in the round-trip suites, since its purpose is what Word puts on the clipboard for it |
| `fonts-embedded.docx` | `build/fonts-embedded.mjs` built the source with the engine (Noto Sans, OFL, on regular, bold, italic and bold-italic paragraphs beside the default font); Jason resaved it in Word 2010 on the reference machine with the entire font embedded, four obfuscated `.odttf` parts (ED-003 section 8.4, step F3) |
| `fonts-embedded-subset.docx` | the same script's second document, with 13 distinct characters of Noto Sans over the four faces (Word subsets only a font used for fewer than 32; the first attempt, over pangrams, came back whole); Jason resaved it in Word 2010 with "embed only the characters used", `w:subsetted="1"` on each face. Its `cmap` format 12 subtables carry a non-zero language field that browsers refuse, which the model repairs on the copy (`docs/developer/upstream-limitations.md` item 5; ED-003 section 11.18 items 4 and 5) |
| `fonts-embedded-subset-m365.docx` | the same source resaved by Jason in Microsoft 365 Word (16.0) with "embed only the characters used", then `build/fonts-embedded-strip.mjs` removed the Aptos and Aptos Display embeddings (Microsoft's fonts, embedded beside Noto Sans; the `w:font` entries stay, the parts and relationships go) so the fixture is redistributable. Its Noto Sans italics carry the same non-zero cmap language field as Word 2010's; its regular and bold are Noto Sans 2.001, which has no format 12 subtable, and need nothing (ED-003 section 11.18 item 5). Not Word's bytes in the font table and the `.rels` parts, which the engine rewrote |
| `opendope/binding-simple.docx`, `opendope/conventions.docx`, `opendope/non-element-bind.docx`, `opendope/hyperlink-binding-test.docx`, `opendope/picture.docx`, `opendope/invoice_Saxon_XPath2.docx` | `docx4j/docx4j-samples-docx4j/sample-docs/databinding` | OpenDoPE templates (E2.b, ED-003 section 4.8): plain Word bindings with no OpenDoPE parts (data part rule 2), six candidate data parts (rule 4, REQ-003), an extended bind (`count()`), hyperlink values, a picture binding, and an XPath 2 template whose condition needs the `xpath2` engine |

`opendope/` also holds the data files the Java oracle pairs with the templates (`manifest.json`
names the pairs; `invoice.docx` and `invoice2013.docx` are the root fixtures), and `goldens/`
once core-ts's `Docx4J.bind` harness has produced them. `opendope/refresh.mjs` copies all of it
from the sibling checkouts; `test/oracle.test.mts` skips with a notice for a pair without a golden.

`build/` holds the scripts that make the built fixtures with the engine (ED-003 section 10); run one
from `packages/editor-model` and commit its output, so the suites read committed documents.

`large/anon-250.docx` is the benchmark document (ED-002 section 7.5), and the large real-world document a few
tests open: an anonymised copy (2026-11-17) of `ms-docx-250.docx`, the [MS-DOCX] Open Specification
concatenated with itself, which was Microsoft's and so could not ship. Its markup is that document's, with every
character of text replaced by another (lengths kept; the words are Latin-like, with no English), every external
link `http://example.invalid/n`, each picture a 2 by 2 PNG, and the core properties emptied. It was made by docx4j's
Anonymiser, which also dropped the sections' three `w:footerReference`s; they were put back as the original had them
(`rId324`, `rId629`, `rId630`, `default`), so the footers show at the page boundaries (ED-002 section 7.5). 255 pages, 9,178 paragraphs, 68 tables (46 page-break
markers in cells), three footers, no footnotes. The per-fixture suites read this directory's root only, so it runs
in the benchmark and the tests that name it alone.

Still wanted (ED-002 section 3.8): a math zone and an `altChunk` for the opaque report.

- `opendope/xpaths-qualified-attributes.docx`: core-ts's check 38 handler template (its `test/fixtures/check38/38b-handler-template.docx`), whose XPaths entries carry `w:`-qualified binding attributes where the schema's are unqualified (ED-005 section 12.87), for `test/opendope.test.mts`.
- `pictures/`: core-ts check 38's picture files (copied from docx4j-core-ts's `test/fixtures/check38/`): `38a-picture-template.docx` (Word's picture control mapped to a node holding an image, showing its placeholder), `38a-picture-bound.docx` and `38b-handler-bound.docx` (the engine's bound saves), for `test/stale.test.mts` and `e2e/pictures.spec.ts` (ED-005 section 12.89).
- `theme-colour-cache.docx`: core-ts's `test/fixtures/check53/53-reads.docx` (check 53, Word 365): its heading styles carry a theme colour with a stale cached `w:val` (Heading 1 `365F91` cached, `0F4761` drawn), pinning that the stylesheet draws a colour through the engine's resolver (ED-008 section 8.8 item 13; core-ts 0.6.0, CR-008 phase 2).
- `theme-colour-run.docx`, `normal-form-values.docx`: built by `build/theme-colour-run.mjs` on `odd-even-first-page.docx` (ED-008 section 8.8). The first holds runs whose `w:val` disagrees with the theme reference beside it, in a paragraph, with a tint, in a tracked insertion and deletion, a table cell, an inline control and the headers, and one whose value agrees. The second holds values a save writes as read and normal form would change: a lower-case colour, a named colour, a highlight outside Word's list, and one in each header.
- `authorless-contents.docx`: built by `build/authorless-contents.mjs` (ED-005 section 12.113): what ONLYOFFICE 9.4 leaves when it refreshes a Contents with tracking on, in its shape with text of our own, the file it was seen in being private. Contents lines of one `w:ins` per run with no author, inside a hyperlink, their marks `w:author=""`; three `w:ins` around empty runs beside each heading's bookmark; a named reviewer's real changes, one with an empty wrapper of no author inside it.
- `mc/mc-block.docx`: docx4j's `mc-fun/mc-block-newest-first.docx` (VERSION_17_3_4, `1cff2cd5c`; core-ts `test/mc-fun/`): a block-level `mc:AlternateContent` whose selected branch holds a paragraph, addressed `body/2/0` by the engine since core-ts 0.6.0 (CR-013); the companion lists it and refuses to edit it (ED-009 section 9.5).
- `mc/mc-simple.docx`: Jason's probe of which branch a reader takes (Word 365, 2026-10-20; the share's `Office 2016/mc-fun/`, for the text extraction paper's synthetic mc cases): a run whose only item is an `mc:AlternateContent` with a `wps` Choice holding `The Earth is flat` and a Fallback holding `The Earth is round` as bare `w:t`. `mc/mc-branches.docx`: built from it by `build/mc-branches.mjs` for ED-009 (the body replaced by its eight cases: text branches; a tab and a break in the Choice; an unknown Choice with a Fallback; one without; text and a picture in the Choice; Jason's ten Choices; the ten with `w16du` first; the paragraph-level shape with runs; added 2026-10-20 for core-ts CR-010 section 15, the ten with `w16sdtdh` first and with `w16sdtfl` first, cases 9 and 10, both prefixes understood by Word 365 as Jason read them the same day; and since core-ts CR-010 section 15 made `x14ac` an understood prefix, cases 11 and 12 on a private prefix `zz9`, with and without a Fallback, the unknown-Choice cases). For `test/alternate-content.test.mts`, `test/concurrency.test.mts` and `e2e/alternate-content.spec.ts`. `mc/mc-for-word-word365.docx`: Jason's Word 365 save (2026-10-20) of `test-results/mc-for-word.docx`, built by `build/mc-for-word.mjs` from `mc-branches.docx` (ED-009 section 4's four shapes; the tracked deletion accepted in Word): every element resolved to plain runs, every Fallback dropped (ED-009 section 9.2; `test/alternate-content.test.mts`). `mc/mc-branches-check49-word365.docx`: Jason's Word 365 save of `mc-branches.docx` after check 49's Script Lab snippet (`Office 2016/mc-fun/check49-script-lab.js` on the share) read every paragraph's text, searched for Choice-only and Fallback-only words, inserted "YY" after the first "flat" hit through its Range and appended its report to the document (ED-009 section 9.2 item 4). None read by `allFixtures`.
- `tables/table-fit-fixed.docx`, `tables/table-fit-auto.docx`: ED-007's five tables (section 5; built by `build/table-fit.mjs` with the engine, 2026-10-14, E added 2026-10-15), with fixed widths and automatic: A three equal columns with one long, B a cell merged over two columns and a vertical merge, C a row of exact height, one of minimum height taller than its content, and an identifier that cannot break (the fit releases both heights), D five columns with one long, E ten narrow columns under long product names whose longest words together are wider than the table (section 8.1 item 12; the names and prices fictional since 2026-11-19, of the same lengths and break points); for `test/tables.test.mts` (`table.columnWidths`) and `e2e/tablefit.spec.ts`. Not read by `allFixtures`.
- `tables/check44/`: editor check 44 (ED-007 section 8.1 item 13): `optional-breaks-probe.docx`, built by `build/optional-breaks-probe.mjs`, a word in a cell too narrow for it, plain, with U+200B, with `w:softHyphen` and with a hyphen, and `optional-breaks-probe-Word365.docx`, Jason's Word 365 save (both kept; Word breaks at either), for `test/check44.test.mts`. The words were replaced 2026-11-19 with fictional ones of the same lengths and break points, the Word save edited in place: Word measured the earlier words.
- `tables/check42/`: editor check 42 (ED-005 section 12.83 items 5 and 6): `insert-table-probe.docx`, built by `build/insert-table-probe.mjs`, and `insert-table-word365.docx`, Jason's Word 365 save after Insert Table under each heading, for `test/check42.test.mts`.
- `tables/check41/`: core-ts check 41 (ED-005 section 12.82 item 6; copied from docx4j-core-ts's `test/fixtures/check41/`): a two-column table with shading, borders, vertical alignment and paragraph formatting per column, with Jason's Word 365 saves after Insert Right and Insert Left, for `test/check41.test.mts`.
- `tables/check40/`: core-ts check 40 (ED-005 section 12.82; copied from docx4j-core-ts's `test/fixtures/check40/`): the window table with unequal columns (`40a1`, `40a2`), the table with a spanning cell (`40b1`, `40b2`), each with Jason's Word 365 save after Insert Right or Insert Left, for `test/check40.test.mts`.
- `tables/check39/`: editor check 39 (ED-005 section 12.79): `table-autofit-probe.docx`, built by `build/table-autofit-probe.mjs` (a two-column table in each layout mode), and `table-autofit-probe-word365.docx`, Jason's Word 365 save after Insert Right in each, for `test/check39.test.mts`.
- `tables/`: four of core-ts's table probes (`tables-conditions.docx` and `tables-conditions-mode12.docx`, made by core-ts's `scripts/make-table-conditions-fixture.mjs` and held to docx4j, not Word; `tables-banding-merged-row.docx` and `tables-banding-band-size.docx`, docx4j CR-030's probes T9 and T10), copied from docx4j-core-ts's `test/fixtures/parity/` for `test/tableconditions.test.mts` and `e2e/table-conditions.spec.ts` (ED-005 section 12.65). Apache-2.0, as docx4j's.
- `tables/cell-last-block-10-only-sdt-p.docx`, `tables/cell-last-block-12-only-sdt-tbl.docx` and its `-repaired.docx` (ED-005 section 12.109; docx4j-25's probes of 2026-10-22, from the share's `Office 2016/cell-last-block/`, copied unedited): a cell holding only a block-level control with a paragraph (10, which Word opens), and one holding only a control that ends in a nested table (12, unreadable), with Jason's Word 365 repair (an empty `w:p` appended to the cell after the control), for `test/tables.test.mts`. Not read by `allFixtures`.

### `clipboard/` (ED-010 section 9.1, not read by `allFixtures`)

- `clipboard-probe.docx` — built by `build/clipboard-probe.mjs`: a comment (whole and part), another author's
  insertion and deletion, a bookmark, a picture, a footnote, a numbered item and a heading, with nine target
  paragraphs for Word's own copy and paste.
- `clipboard-probe-word365-off.docx` — Jason's Word 365 run with tracking off (2026-10-22): eight pastes.
- `clipboard-probe-word365-on.docx` — the same with tracking on as Author A: the revisioned paragraph pasted.
- `clipboard-cut-word365-off.docx`, `-on.docx`, `-on-accepted.docx` — Jason's Word 365 run of 2026-10-29 (ED-010
  section 9.3): case 5's bookmarked paragraph cut whole and pasted over T5. Tracking off: the bookmark goes with the
  paragraph. Tracking on, as Author A: a move (`moveFrom` at the original place, the bookmark inside it; `moveTo` at the
  paste, no bookmark; T5 a deletion), and after Accept All the bookmark is gone.
- `dup-docpr-2.docx`, `dup-bookmark-2.docx` — two inline drawings sharing a `wp:docPr` id and name; two bookmarks
  sharing a name and id. `-word365.docx` beside each: Word 365 opened both clean (no repair) and on save renumbered
  the second drawing and dropped the first bookmark, keeping the second.

### `br-placement/` (ED-008 section 8.8 install item 17, not read by `allFixtures`)

docx4j's `docx4j-core-tests/src/test/resources/br-placement/` (docx4j CR-073; its README there says what each case holds and
where the load survey found it): thirteen one-case files, each a `w:br` outside a `w:r` (directly in a paragraph, a
hyperlink, a cell, a footer, the body; inside a `w:t`, a nested `w:t`, a `w:instrText`), built by the two Python scripts
there, and beside each `-word365.docx`, Word 365's save of it (Jason's run, 3 November 2026). Copied 2026-11-10 at the
core-ts 0.7.0 install, which reads them as Word does (CR-021). For `test/br-placement.test.mts`.

### `xhtml/` (ED-005 section 12.117, not read by `allFixtures`)

`check37a-template.docx` is core-ts `test/fixtures/check37/37a-template.docx` (Word check 37, built by that folder's
`build.mjs`; copied 2026-11-11): a block-level and a run-level rich text control tagged `od:ContentType=application/xhtml+xml`
and bound through `od:xpath` entries, never filled, and a plain text control mapped to the same node. The data part holds a
heading, bold, a link, a two-item list and a table for the first, two paragraphs for the second. For the Preview tab's
conversion: `test/preview.test.mts` and `e2e/template.spec.ts`.

### `fragments/` (ED-005 proposal 46, not read by `allFixtures`)

The files for proposal 46's Word check, part 6(a): Word fragments (`od:progid=Word.Document`). Built by hand as XML by
`build/fragments-for-word.mjs`, and run in Word 365 through the portfolio's Word bridge. `fragments/README.md` says what
each holds and what to read from Word's saves, which go beside them. `test/fragments.test.mts` and `e2e/fragments.spec.ts`
read 46a's fragment and 46c's template.

### `pictures-dataurl/` (ED-005 section 12.125, not read by `allFixtures`)

The files for section 12.125's Word check: a picture control mapped to a node holding a data URL, where Word writes bare base64.
Built by hand as XML by `build/picture-dataurl-for-word.mjs`, and run in Word 365 through the portfolio's Word bridge.
`pictures-dataurl/README.md` says what each holds and what to read from Word's saves, which go beside them.

### `table-indent/` (ED-005 section 12.132, not read by `allFixtures`)

The files for section 12.132 item 5's Word check: where Word puts a left-aligned table's edge, and its cells' text, against
the body text, in compatibility modes 15 and 14 and with no mode stated; seven tables varying `w:tblInd`, the table's left
cell margin and the first cell's own. Built by `build/table-indent-for-word.mjs`, by hand as XML. To be run in Word 365
through the portfolio's Word bridge. `table-indent/README.md` says what each rule predicts and what to read from Word's
saves, which go beside them.

### `row-shading/` (ED-005 section 12.128, not read by `allFixtures`)

The files for section 12.128's Word check: a table cell's own shading in the forms a "no colour" could take, under a style
that shades the cell; a vertical merge's shading on one `w:tc` or on all of them; and a file the editor's `row.shading` wrote.
Built by `build/row-shading-for-word.mjs`, by hand as XML except `shd3-editor.docx`, the model's own save. To be run in Word
365 through the portfolio's Word bridge. `row-shading/README.md` says what each holds and what to read from Word's saves,
which go beside them.

### `proofing/` (ED-005 section 12.120, not read by `allFixtures`)

`proofing.docx`, built by hand by `build/proofing.mjs`: a document in Australian English with a French run, a run marked
`w:noProof`, a paragraph in a style with `w:noProof`, Japanese text whose `w:lang` names `ja-JP` for East Asian text, and a
character style in German. For `e2e/proofing.spec.ts`.


### `text-extraction/` (ED-008 step 2, not read by `allFixtures`; not in the public repository)

The public repository leaves this folder out (the release plan's "what crosses the fence", Jason, 2026-11-18). Its two
tests load the runner through `test/text-extraction-suite.mts` and skip where the folder is absent, so the public suite
runs without it.

The text extraction specification's conformance suite, unedited, with `manifest.json` naming its sources and each file's
SHA-1 (`test/text-extraction.test.mts` checks them; recopy with `node scripts/copy-text-extraction-suite.mjs <docx4j>
<commit> <core-ts> <tag>`, never edit):

- `suite/` is docx4j's `docx4j-core-tests/src/test/resources/text-extraction/` at `64fae4cd0` (the docx4j 17.3.5 release,
  2026-11-01), the commit core-ts 0.8.0's own copy is at: the probe's directory `phase0/` whole (`extraction-probe.docx`,
  cases 1 to 49; `extraction-probe-external.docm`, 50 to 54; `extraction-probe-structure.docx`, 55 to 57; the inspection
  and HTML probes; `expected-tokens.tsv`, `expected-findings.tsv`, `expected-order.tsv`, `case-statements.tsv`,
  `goldens.tsv` and the twelve goldens; the inspection tables; Word's evidence and the baseline), and the suite's
  `README.md`, `GOLDENS.md` (the format, REQ-1201's normalisation, the adapter contract) and `INSPECTION.md`;
- `conformance.mjs` is core-ts's runner at its tag `0.8.0` (`test/text-extraction/conformance.mjs`, its mirror of docx4j's
  `ExtractionConformance`). Its types are the editor's own, in `test/text-extraction-suite.mts`, outside the folder.

For `test/text-extraction.test.mts` (the installed engine held to the suite, with the release's own counts) and the
companion's `test/tools.test.mts` (each reading tool held to the engine's output and to the goldens). Copy again at a
core-ts install whose release recopied its suite, from the commit its manifest names.

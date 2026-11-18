# The agent's tool contract (`docx4j-editor-mcp`)

**Status: implemented** in E4.f step F3, 2026-09-29 (ED-005 section 12.27), each tool held to it by
`packages/editor-mcp/test/tools.test.mts`. Cleared for F3 by ED-005's second review point (section
2.1) the same day, with the review's findings addressed (ED-005 section 12.25); written in step F2.
What building it settled or added is marked "(F3)" where it stands.
It answers, tool by tool, the questions the portfolio's vocabulary asks of every shared operation
(`docx4j-portfolio/docs/mcp_tool_vocabulary.md` section 7): input schema, addressing, mutation,
output, errors, validation, tracked-change behaviour, atomicity. The tool names are the
vocabulary's (ED-005 proposal 11). Where a shared tool's parameters are also the Python server's
(`docx4j-python` CR-006), this contract takes CR-006's names; what differs is listed in section 5,
for the vocabulary.

## 1. What the server is

`docx4j-editor-mcp` puts an agent in a co-editing room beside a person. The person sees the agent in
the People list and every edit it makes as a tracked change under its name, waiting for accept or
reject. Four rules follow, and every tool keeps them:

1. **Always tracked.** Every change the agent makes to the document's content is a tracked change by
   the agent (`--name`, default "Claude"), whatever the document's own setting, which the agent never
   changes (ED-005 section 8.5). **What is not tracked**, because Word has no revision for it, and
   what a person therefore cannot reject:
   - a comment, its reply and its resolved state (Word never tracks comments);
   - a style added to the styles part, as when a Markdown heading names a style the document lacks,
     or a comment is the document's first;
   - a numbering definition made for a Markdown list;
   - a hyperlink's relationship;
   - a content control wrapped or unwrapped, its properties, and a binding's write to the custom XML
     and OpenDoPE parts (`insert_content_control`, `remove_content_control`, `set_content_control`,
     `bind`): the model holds no revision for a control (ED-005 proposal 5), though the content
     inside one is tracked like any other.

   The change report lists each such part under `parts` (section 2.6).
2. **One call, one change.** A call's edit reaches the room as one transaction. Everything a call
   needs is checked before anything is written; a call that fails writes nothing.
3. **The person is typing.** Addresses and text the agent read can be stale by its next call.
   `if_text` (section 2.4) and the change report (section 2.6) are how the agent keeps up.
4. **No scripts, no raw parts.** No tool runs a script or a macro, and no tool edits a package part
   as XML; `insert_ooxml` inserts content, which is tracked like any other.

## 2. What every tool shares

### 2.1 The document handle

Every tool but `open` and `list_documents` takes **`document`**: the room id `open` returned (or
`list_documents` lists). MCP 2026-07-28 wants explicit handles; an open stdio process is not a
session. An unknown handle is `document.not_open`.

### 2.2 A paragraph's text

**A reading tool shows the text in the view it is asked for** (section 2.13; `accepted` by default), **and every text
this contract compares or counts is the paragraph's text as accepted:** text inserted as
a tracked change is in, the agent's own pending insertions included; text deleted as a tracked change
is out, and so is hidden text and a content control's placeholder (a control showing its placeholder
reads as empty). The characters are Office JS's, Word's own stream (core-ts 0.6.0): a tab `\t`; a line
break U+000B, a page break U+000C, a column break U+000E; a footnote or endnote reference U+0002; a
non-breaking hyphen U+001E and an optional hyphen U+001F; any symbol character "("; an equation its
text; a picture, a field character, a bookmark or comment marker no character. This is the engine's
`Paragraph.text`. `find`'s plain query takes Word's Find codes (`^t` a tab, `^l` a line break, `^m` a
page break, `^f` and `^e` a note reference, `^~` and `^-` the hyphens, `^g` a picture, `^^` a caret),
passes over an optional hyphen unless the query names it, and does not cross a picture.

It is what `outline`, `find`, `read` and `extract_text` show by default; what a span's `offset` counts (section 2.3); what
`if_text` is compared against (section 2.4); and what a change report's `text` fields carry
(section 2.6). `get_tracked_changes` says what is pending.

### 2.3 Addresses and spans

A tool's **`address`** names a block, in one of the vocabulary's three forms, as the engine resolves
them (`@docx4j/core-ts` `Address`):

| Form | Example | Stable while |
|---|---|---|
| Ordinal | `"body/3"`; `"body/7/0/1/0"`, table 7's row 0, cell 1, paragraph 0; `"body/3/textbox:0/1"`, the second paragraph of the first text box anchored in paragraph 3; `"body/2/0"`, the first block of a kept `mc:AlternateContent`'s selected branch standing at block level | nothing before it is added or removed |
| Paragraph id | `{ "paraId": "5A2B1C3D" }`, or the string `"w14:5A2B1C3D"` | always, where Word wrote one (not every file has them; a table never) |
| Text | `{ "contains": "Chapter 1" }` | the text stays; the first paragraph containing it, in document order |

**The paragraph id's spelling:** every output field named `paraId` carries the 8 hex digits alone
(`"5A2B1C3D"`). An input accepts that, the prefixed `"w14:5A2B1C3D"`, and either inside
`{ "paraId": … }`, in any case.

A text box's paragraphs are listed after the paragraph (or table) anchoring the box, at `textbox:`
addresses, and are read-only: a mutating tool on one is `target.invalid`. A block in a kept
`mc:AlternateContent`'s selected branch (block, row or cell level) is listed at its address through the
element and is read-only in the same way; a person can edit it in the editor. `find` with an XPath
reports a hit inside a text box once, at the box's own address, whichever branch of the box's
`mc:AlternateContent` the XPath selected.

`outline`, `find` and `read` return the ordinal and, where there is one, the paragraph id. **Prefer
the paragraph id**: the person's edits shift ordinals. An address that resolves to nothing is
`address.not_found`, carrying the nearest address that exists (`nearest`) and the hint to call
`outline`. A text address is always the first match; the change report names which paragraph that
was. A block the agent deletes stays in the document, deleted, until the person resolves the
deletion, and keeps its ordinal: `edit`'s Replace leaves the old block where it was and puts the new
ones after it, which the report's `created` addresses give.

A tool that acts on part of a paragraph (`set_format`, `delete`, `comment`) takes a **span** beside
the `address`:

- nothing more: the whole paragraph;
- **`anchor`**, the words: their first occurrence in the paragraph, or the `occurrence`-th
  (1-based);
- or **`offset`** and **`length`**, in the paragraph's text (section 2.2).

A span that does not resolve (the words are not there, the offsets are past the end) is
`range.invalid`.

### 2.4 `if_text`

A mutating tool takes an optional **`if_text`** for its target: the paragraph's text (section 2.2)
as the agent last read it, under the `accepted` view, in either spelling: Office JS's (`Paragraph.text`, what a change
report and `target.stale` carry) or the text model's (what `outline`, `find` and `read` show: a note's reference no
character, a line break a newline; ED-008). A text read under `original` or `markup` is not the accepted text, and is
refused as stale. A block `outline` lists as joined (section 2.13) is several paragraphs: `if_text` is one paragraph's
own. When the paragraph's text now differs, the call writes nothing and fails
with `target.stale`, carrying the paragraph's current text, address and paragraph id, so the agent
can decide again. It is per target, not per document, because a person typing elsewhere must not
refuse the agent's call (vocabulary section 7). A tool whose target is not one paragraph
(`replace_text` over the document or a table) refuses an `if_text` with `arguments.invalid`.

### 2.5 Locations, and editing at block level

Office JS's locations, as the vocabulary fixes them: `Replace`, `Before`, `After`, `Start`, `End`.
**Editing is at the level of blocks.** `edit`, `insert`, `delete` (without a span) and `insert_ooxml`
replace, add or remove whole paragraphs and tables. `Start` and `End` mean the start and end of a
container: the body when `address` is omitted, a table cell, or a content control. Text inside a
paragraph changes through `replace_text` (with the paragraph as its `scope`), `set_format`, and
`delete` with a span. Section 3.3 lists each tool's locations. A location a tool does not take is
`location.invalid`.

### 2.6 Markdown

**In** (`edit`, `insert`, and a comment's `text`), a subset, each mapped as Word would make it:

| Markdown | Becomes |
|---|---|
| a paragraph | a paragraph in the target's style: a replaced paragraph's own; before or after one, the style Word's Enter gives after it (its style's `w:next`), else its own; Normal at a container's start or end (F3) |
| `#` to `######` | Heading 1 to 6, by the document's style names; a heading style the document lacks is added from docx4j's defaults, Heading 5 and 6 from Word's definitions (`knownstyles.mts`, UPSTREAM-20) (untracked, rule 1) |
| `**bold**`, `*italic*`, `~~strike~~` | direct formatting on the runs |
| `- item`, `1. item` (nested by indentation) | bulleted and numbered list items through the document's numbering (a new definition untracked, rule 1), as the editor's list commands write them (`w:numPr` alone); a second numbered list in one call starts again at 1 (F3) |
| a GFM table | a table, the first row its header |
| `[text](https://…)` | a hyperlink (its relationship untracked, rule 1) |
| anything else (images, HTML, code blocks, inline code, block quotes, rules, footnotes, reference-style links) | refused with `markdown.unsupported`, naming the construct; `insert_ooxml` is the escape hatch |

A link's address is `http`, `https` or `mailto`. A comment's `text` takes paragraphs with bold,
italic and strikethrough only: a heading, a list, a table or a link in a comment is
`markdown.unsupported` (F3).

`edit` also takes plain **`text`** in place of `markdown`, as CR-006's does.

**Out** (`read`): the engine's Markdown of the text model (core-ts's `MarkdownExporter`, docx4j's byte for byte;
ED-008 proposal P3), each block preceded by an HTML comment naming its address, `<!-- body/3 paraId=5A2B1C3D -->`, and
nothing else added: removing those comments gives the engine's own Markdown of the document. It is CommonMark with GFM's
tables and strikethrough: a heading from its outline level, a list from its numbering, a table's cells' paragraphs `<br>`
apart and its merges as GFM can say them, notes as `[^1]` with their definitions after the body, pictures with their
description, CriticMarkup under `markup` (section 2.13). It is a wider set than Markdown in takes: what `read` writes
cannot always be written back through `edit`. The result's structured content carries the same blocks as
`{ address, paraId?, style, kind, markdown }`. `get_comments` writes a comment's text as Markdown paragraphs, escaped.

### 2.7 The change report

Every mutating tool returns (vocabulary rule 6), as structured content and as a one-line summary:

```json
{
  "summary": "Replaced \"30 days\" with \"45 days\" in 2 paragraphs",
  "touched": [{ "address": "body/12", "paraId": "5A2B1C3D", "text": "…the paragraph's text now…" }],
  "created": [{ "address": "body/13", "text": "…" }],
  "removed": [{ "address": "body/14" }],
  "revisions": [{ "kind": "insertion", "author": "Claude", "address": "body/12", "text": "45 days" }],
  "parts": ["/word/styles.xml"],
  "warnings": ["the text address matched body/12; later matches were not changed"]
}
```

`touched` and `created` carry each paragraph's new text (section 2.2), which is the agent's next
`if_text`; a created table's `text` is its cells' text, a line each. Ordinals in a report are as of
the moment the call's change was made. `parts` names the package parts the call changed beside the
document's content, the untracked ones of rule 1.

`revisions` lists the call's own changes as Office JS lists changes (F3): a paragraph mark is `\r`,
and touching insertions by one author across paragraphs are one change, at the paragraph it starts
in. So a Markdown payload of three paragraphs is one insertion whose text holds two marks.

### 2.8 Errors

A failed call is an MCP tool error (`isError: true`) whose structured content is
`{ "code", "message", "hint", … }`. The codes are dotted, as docx4j-python CR-003's are, so that an
agent recovers the same way whichever docx4j server answered. Where the meaning is CR-003's, so is
the code; the editor's own are coined in the same form:

| Code | Also CR-003's | When | Carries |
|---|---|---|---|
| `address.not_found` | yes | an address resolves to nothing | `address`, `nearest` |
| `range.invalid` | yes | a span does not resolve in its paragraph | `address`, `anchor` or `offset` |
| `target.invalid` | yes | an operation the target does not allow (Markdown into a table cell's structure, a table where a paragraph is needed) | `address` |
| `location.invalid` | yes | a location the tool does not take | `location`, `allowed` |
| `style.not_found` | yes | a style name or id neither in the document nor in docx4j's defaults | `style` |
| `document.not_open` | | the `document` handle is unknown, or the room was closed | |
| `target.stale` | | `if_text` differs from the paragraph's text; or the person's typing reached a paragraph the call was changing while it ran (section 2.10) | `current: { address, paraId?, text }` (`address` absent when the paragraph cannot be found again; `null` when the person deleted it) |
| `markdown.unsupported` | | Markdown the subset does not map | `construct` |
| `ooxml.invalid` | | OOXML that does not parse, or names a relationship it does not bring | |
| `path.not_found` | | a path under a root that names no file (`open`), or no folder (`output_path`) (F3) | |
| `arguments.invalid` | | arguments the schema allows but the tool cannot take together (an `if_text` on a scope wider than a paragraph); or a measurement past the range Word stores or takes (`left_indent: 400000`), refused by the engine before anything changes, its message naming the range (core-ts 0.7.0, CR-002 section 65); or a reading tool's option it does not take, or one the engine does not implement (section 2.13, ED-008 P5) | `argument`; for a text-model option, `statement`, the specification's (`REQ-506`) |
| `path.not_allowed` | | a path outside every `--root` | `path` |
| `path.exists` | | `output_path` names an existing file other than the document's own, without `overwrite: true` | `path` |
| `tracking.not_allowed` | | `track_changes` asked to turn tracking off | |
| `resolve.not_allowed` | | accept or reject without `--allow-resolve` | |
| `agent.stopped` | | the person pressed Stop for the agent in this document: a call on its room, or `open` of its file or a copy the agent saved from it | |
| `relay.down` | | the companion's relay stopped and could not be restarted (it restarts a stopped relay on its own port first, and a call waits for that): `open`, an invitation and every edit are refused until the companion is restarted, since nothing would reach a room; reads answer from the agent's own copy | |
| `internal.error` | | the companion failed, not the call (F3) | |

`bind` reports CR-003's `binding.*` codes where they apply: `binding.no_repeat_node` (a repeating
section's XPath selects nothing) and `binding.document_element` (the XPath selects the data's document
element); and one coined in the same form, `binding.no_node`, an XPath that selects no node in the data
part, where Word writes nothing (F3). An XPath that does not parse is `arguments.invalid`.

### 2.9 The inline cap

As vocabulary section 6 requires of every server: a result past the inline cap
(`--max-inline-chars`, default 200,000 characters) is written to `output_path` when the call gives
one, and otherwise truncated, with `truncated: true` and `chars` (the untruncated length) in the
result. `outline`, `read` and `extract_text` take `max_chars` to set a lower cap for one call.

### 2.10 Concurrency

Calls to one room run one at a time, in order. Each call's change is made against the document as it
is when the call runs, the person's latest edits included, and reaches the room as one transaction.
So a person accepting or rejecting a change races the agent's edit only for the relay's latency.
A deletion the agent makes at the moment a person deletes overlapping text keeps both since core-ts
0.3.0 (ED-005 section 12.26, room item (a)).

**The person typing while a call runs** (ED-005 proposal 42). What the person types while a call is
running reaches the companion before the call's change is merged. If it reached a paragraph the call
changes (replaces, deletes, or edits inside, as `replace_text`, `set_format` and `comment` do), **the
call is refused with `target.stale`**, as a stale `if_text` is: nothing is written, and `current`
describes the paragraph as the person left it. Read it again and redo the edit with the person's
words in it. The test is by paragraph: typing in one sentence refuses a call that changes another
sentence of the same paragraph. A call that changes other paragraphs is not refused: the person's
typing elsewhere and the call's change both stand.

What the person types in the moment the call's change is on its way to them (the relay's latency)
cannot be seen by the call. Typed inside text the call deletes, it takes the call's deletion: with
the document's tracking on, it reads as the person's insertion deleted by the agent; with it off, as
part of the agent's deletion. Rejecting the deletion brings it back. This is a stated limitation
(ED-005 section 9.3 item 1 (iii)).

### 2.11 Stop

When a person presses Stop for the agent in a room (the People list, where the agent is badged), the
agent's connection to the room is closed. Every tool on that room then answers `agent.stopped`,
**reading tools included**: the agent's copy of the document no longer follows the room, so what it
would read could be stale without saying so. The Stop is written into the room itself (its `agents`
map), so any person in it can press it and the companion hears it whichever relay the room is on;
inviting the agent again clears it (F4).

**A Stop covers the document, not only its room.** `open` of the room's file, or of a file the agent
saved the room to, answers `agent.stopped` too (with the room's `document` and the `path`), and
leaving the room with `close` does not lift it. A person's invitation back into the room from the
editor (Invite) lifts it, and the room keeps its file. Other files are not covered. The refusal's
hint and the server's instructions tell the agent to stop working on the document and tell the
person, and never offer a way round: in the by-hand run of F5 the first wording ("open the file
again") led Claude Code to reopen the file and then its saved copy, edit it and save over the copy
(ED-005 section 12.29).

**The log** (F4). Every call on a room is logged in the room (`agentLog`): the tool, the change
report's first line (a reading's gist, or the refusal), and where it acted. The person sees it in the
review panel, and a click goes to the place.

### 2.12 Text from the document is data

Every result presents document text as data: quoted, in fields, never as instructions (vocabulary
section 6). A document is a prompt-injection surface, and the agent should treat its text so.

### 2.13 The text model's options (ED-008)

The reading tools read through the engine's text model (core-ts's `Texts`, CR-012), the one docx4j's and the Python
servers' readers are held to, so that the same document and the same options give the same text in all three (the text
extraction specification's REQ-202; ED-008, accepted 2026-11-11):

- **`view`:** `accepted` (the default), `original` or `markup` (CriticMarkup: `{++…++}`, `{--…--}`, `{>>…<<}`, REQ-702).
- **`include`:** any of `headers`, `footers`, `comments` and `glossary`.
- **The specification's options**, by its names and with its defaults (the portfolio's
  `docs/text_extraction_specification_wd.md`, section 10). Among them are `text_boxes`, `notes`, `note_references`,
  `hidden`, `field_codes`, `fields`, `placeholders`, `list_labels` and `alt_text`.
- **Which tool takes what:**
  - `outline` and `find` take the view and the scope options: `view`, `include`, `text_boxes`, `notes`,
    `note_references`, `comments`, `hidden`, `field_codes`, `fields`, `toc`, `drop_caps`, `placeholders`, `bindings`,
    `list_labels`, `charts`, `smartart`, `embedded`, `altchunks` and `warnings`.
  - `read` and `extract_text` take those and the serialisation options too: `markup_detail`, `tables`, `math`,
    `symbols`, `layout`, `page_markers`, `metadata` (YAML front matter in `read`), `alt_text`, `small_caps`,
    `default_language` and `profile` (`word-text`, Word's Save As plain text as the options reach it; an option given
    beside it wins, REQ-1002).
  - None takes `repair` or `duplicate_entries`: they are how a file is loaded, and `open` reads a file Word repairs as
    Word does (section 3.1).
- **What the reading covers:** the room's document as it is when the call runs, the person's typing included, and the
  notes, headers and comments the room's sessions have edited.
- **`find` gives `offset` and `length` under `accepted` alone,** the text a span counts in (P2). A hit under another view has
  its address and snippet only, and the result's `spans` says so. Under `accepted` a hit has no span where its text is not
  one paragraph's text as the edit tools count it: in a note, a header or a comment (`story` names it), across a note's
  reference (which the text model's text holds as no character), or across a paragraph mark `accepted` joins. A hit's
  `address` is the paragraph its words are in.
- **A paragraph joined across a deleted mark** (`accepted`; an inserted one in `original`) is one block at its first
  paragraph's address and paragraph id (REQ-405), `joined` naming the addresses of the paragraphs after it. The edit tools
  take each paragraph at its own address.
- **An option a tool does not take, or a value the engine does not implement** (`fields: update`, `embedded: true`,
  `math: unicodemath`, `symbols: unicode` at core-ts 0.8.0), **is `arguments.invalid`,** with the option as `argument`
  and the specification's statement as `statement`, never accepted silently, as docx4j refuses it (P5). So is an argument
  no reading tool takes: a misspelt option is refused, not ignored.
- **`page_markers: true`:** the markers are current only when the file was paginated after its last edit. In the editor
  that is never after a keystroke: Re-paginate.

## 3. The tools

Schemas are given as JSON Schema's essentials. Every mutating tool also takes `dry_run: true`, which
returns the change report without writing, and the `if_text` of section 2.4.

### 3.1 The room

| Tool | Input | Does | Output |
|---|---|---|---|
| `open` | `path` (under a `--root`) | Opens the file into a new room, the agent joined, and returns the link a person opens to join in the editor; a file already open returns its room. A file Word opens only through a repair (two entries of one name, several main parts, a relationships part in a wrong namespace) is read as Word repairs it, the engine's message for each repair in `repaired`, and a save writes the repaired file (core-ts 0.7.0; Jason, 2026-11-10) | `{ document, link, path, repaired? }` |
| `list_documents` | — | The rooms the agent is in: those it opened, and those a person invited it into from the editor (F4), whose `path` is null | `{ documents: [{ document, path, link, people: [{ name, agent }], stopped? }] }` (F3: structured content is an object) |
| `close` | `document` | Leaves the room; the person's editor stays open | `{ document }` |
| `save` | `document`, `output_path?`, `overwrite?` | Writes the document, every tracked change pending as it is, to its file, or to `output_path` (under a root); a document the agent was invited into has no file of its own, and needs `output_path` (F4) | `{ path, bytes }` |
| `track_changes` | `document`, `mode?` (`TrackAll` \| `Off`) | Reports the agent's tracking and the document's own setting; `mode: "Off"` is `tracking.not_allowed` (the vocabulary's third shape) | `{ mode: "TrackAll", author, document_mode }` |

### 3.2 Reading

| Tool | Input | Output |
|---|---|---|
| `outline` | `document`, `budget?` (default 400 blocks), `max_chars?`, `view?`, `include?` and the scope options (section 2.13) | `{ blocks: [{ address, paraId?, style, text (its first 80 characters), kind: "paragraph" \| "table" \| "chart", rows? (a table's), story? (`text_box`, `footnote`, `endnote`, `header`, `footer`, `comment`, `glossary`; the main story's blocks have none), joined?, hidden?, placeholder?, decorative?, revisions?, numbering? ({ level, format?, value? }), list_label? }], truncated, warnings? }`, the text model's blocks in reading order (a table, then its cells' blocks); a block's `text` is its text in the view, its text runs alone (no note number, no label), and under `markup` its line as `extract_text` writes it |
| `find` | `document`, `text` or `xpath`, `match_case?`, `match_whole_word?`, `match_wildcards?`, `match_regexp?`, `limit?` (100), `context?` (40 characters), `view?`, `include?` and the scope options (section 2.13; with `text` alone) | `{ hits: [{ address, paraId?, offset?, length?, snippet, story? }], total, truncated, spans? }`, the text searched each paragraph block's text in the view, `offset` and `length` under `accepted` alone and only where section 2.13 gives a span, `spans` saying why there are none under another view; an `xpath` is evaluated over a snapshot of the tree as it is (the engine's `bindSnapshot` and `boundsOf`, as `Body.select` is), its hits `{ address, paraId?, kind, snippet }`, the paragraphs and tables it selects or selects inside |
| `read` | `document`, `address?` (the whole document when omitted; a note's, a text box's or a header's address too, as `outline` lists them), `view?` (`accepted`, `original` or `markup`), `include?`, the options of section 2.13, `max_chars?`, `output_path?` | `{ markdown, blocks, warnings, truncated, chars }`: Markdown with addresses (section 2.6), the engine's Markdown of the text model (GFM, REQ-703; CriticMarkup under `markup`); a block at an address keeps its place in the document (a list item its number); `warnings` the model's (REQ-801: `{ kind, address?, selector?, detail? }`) and the editor's (`{ kind, message }`: `address.first_match`, `address.no_text`) |
| `extract_text` | `document`, `address?`, `view?`, `include?`, the options of section 2.13, `max_chars?`, `output_path?` | `{ text, warnings, truncated, chars }`: the engine's plain text of the text model (core-ts's `PlainText`): one line per paragraph (REQ-701), `[n]` at a note's reference and the notes after the body, CriticMarkup under `markup`; the vocabulary's verb, shared by name with docx4j-mcp and the Python server |
| `get_comments` | `document`, `address?` | `{ threads: [{ id, author, date, resolved, address, anchor, text, replies: [...] }] }`, the comment's text as Markdown |
| `get_tracked_changes` | `document`, `address?`, `author?` | `{ changes: [{ key?, kind, author, date, address, text, shown_in_word_only? }] }`, as the review panel lists them (ED-005 section 6.5): a move one entry at its destination; `kind` one of `insertion`, `deletion`, `formatting`, `paragraph_properties`, `move`, `row_insertion`, `row_deletion` (a change of rows alone: rows grouped with the inserted or deleted text touching them are an `insertion` or a `deletion`, as Office JS lists them, core-ts check 31), and, for a revision Word shows that the editor does not draw (`shown_in_word_only`), `table_properties`, `cell_insertion`, `cell_deletion`, `cell_merge`, `section_properties`, `paragraph_mark_formatting`, `numbering`, `equation_insertion`, `equation_deletion` (an equation's text or structure revised in Word's form, inside the equation), `content_control_insertion`, `content_control_deletion` (a control or custom XML element put in or taken whole under tracking, Word's marker pairs around it); no `key` for a cell merge or a list's numbering, which only "all" resolves (F3) |

### 3.3 Editing

| Tool | Input | Locations | Mutation | Tracked as |
|---|---|---|---|---|
| `edit` | `document`, `address?`, `markdown` or `text`, `location` (default `Replace`) | `Replace`, `Before`, `After`; `Start`, `End` of a container | Replaces or adds whole paragraphs at the block | insertions and deletions |
| `insert` | as `edit` | `Before`, `After` (default); `Start`, `End` of a container | `edit` defaulting to `After`; a GFM table payload inserts a table (the vocabulary's `insert_table` folded in) | insertions |
| `delete` | `document`, `address`, a span? | — | The block, or the span's text | deletions |
| `replace_text` | `document`, `find`, `replace`, `scope?` (an address; default the whole document), `match_case?`, `match_whole_word?`, `match_wildcards?`, `match_regexp?` | — | Every match in scope, as Word's Replace All; `$1` references with `match_regexp` | a deletion and an insertion per match |
| `set_style` | `document`, `address`, `style` and/or `alignment`, `left_indent`, `first_line_indent`, `space_before`, `space_after`, `line_spacing` | — | The paragraph's style and properties; a style is looked up by id first, then by name, so an id wins where both match | a paragraph-properties change |
| `set_format` | `document`, `address`, a span?, any of `bold`, `italic`, `underline`, `strike`, `font`, `size`, `color`, `highlight` | — | Character formatting over the span (the whole paragraph without one) | a formatting change |
| `insert_ooxml` | `document`, `address?`, `ooxml` (a `w:p`, `w:tbl` or a flat OPC fragment), `location` | `Replace`, `Before`, `After`; `Start`, `End` of a container | Inserts the content as parsed | insertions |
| `comment` | `document`, `address` and a span? (the words it is on), or `reply_to` (a comment id); `text` (Markdown); `resolve?` | — | A comment on the span, or a reply in the thread; `resolve: true` resolves the thread. A span in a pending insertion stays on its words, the markers placed as Word places them (core-ts 0.3.1) | **not tracked** (rule 1) |

**The vocabulary's worked example** (section 7, `replace_text`):
- every match in scope is replaced, not the first;
- paragraph ids survive, since a replacement changes runs, not paragraphs;
- the report names every touched paragraph with its new text;
- the call is one change, all matches or none;
- a stale paragraph is met by `if_text` when the scope is that paragraph;
- tracking is the same as for every other editing tool.

### 3.4 Content controls

The editor's own (vocabulary section 4: the shape other servers should copy).

| Tool | Input | Does |
|---|---|---|
| `insert_content_control` | `document`, `address` and a span? (or several blocks of one container, from `address` to `to`; or a table), `kind` (`RichText`, `PlainText`, `Checkbox`, `DropDownList`, `Date`, `Picture`, `RepeatingSection`), `tag?`, `title?` | Wraps the target (untracked, rule 1); the report's `control` is `{ id, tag, title, kind, address }` |
| `remove_content_control` | `document`, `control`, `keep_content?` (default true) | Unwraps (untracked, rule 1), or removes the content with it (its content's deletion tracked) |
| `set_content_control` | `document`, `control`, any of `tag`, `title`, `placeholder`, `lock` (`none`, `content`, `control`, `both`), `checked` (a checkbox), `items` (a drop-down list, `[{ text, value? }]`), `date_format` (a date) | Sets properties (untracked, rule 1); a property of another kind is `target.invalid` |
| `bind` | `document`, `control`, `xpath`, `store_item?`, `prefix_mappings?` | Binds to the document's custom XML (OpenDoPE), as the Template perspective does, and shows the node's value in the control at once, as Word does (untracked, rule 1) |

`control` is `{ id }` (its `w:id`), `{ tag }` (the first control with it; a warning when others share
it) or `{ address }` (a block content control's address) (F3).

The content inside a control is tracked like any other; the control itself is not (rule 1).

### 3.5 Resolving (`--allow-resolve` only)

| Tool | Input | Does |
|---|---|---|
| `accept_tracked_changes` | `document`, `key?` or `address?`, `author?` | Accepts the change named, those at the address, or all (the vocabulary's default scope), as the review panel does, through the editor's four commands (F3); `{ summary, resolved }`. A key no longer there is `target.stale` |
| `reject_tracked_changes` | as above | Rejects them |

Without `--allow-resolve` both are listed and answer `resolve.not_allowed`: an agent resolving its own
changes undoes the review the tracking exists for (ED-005 proposal 12).

## 4. What the contract leaves out

- **Scripts and macros**, on any tool (ED-005 proposal 13).
- **Raw part edits.** The parts panel's Apply is a person's (E4.b).
- **Notes, headers and comments' own text as edit targets.** Their addresses (`footnote:2/0`,
  `header:rId3/0`) are listed by `outline` and readable through `read` and `extract_text`, not writable, in E4.
- **A document version.** `if_text` replaces it (section 2.4).

## 5. Where the editor's shapes differ from CR-006's, for the vocabulary

The vocabulary allows a shared name to differ in shape, provided the difference is recorded (its
rule 8). The editor's shared tools differ from CR-006 in these ways, each a consequence of a live room
or of the engine; the vocabulary's section 8 row for the editor should list them (the portfolio's
work):

| Tool | Here | CR-006 | Why |
|---|---|---|---|
| every tool | `document` (a room id) | `path` | a room, not a file, is the handle (vocabulary section 7) |
| mutating tools | `if_text` | a version or hash (vocabulary section 7) | the person is typing |
| `set_format`, `delete`, `comment` | a span: `anchor` and `occurrence`, or `offset` and `length` | `address` (a paragraph), or `text` to anchor a comment on | formatting and deleting part of a paragraph without rewriting it |
| `find`, `replace_text` | `match_wildcards`, `match_regexp` too | `match_case`, `match_whole_word` | the engine's search has them (ED-005 section 3.2) |
| `track_changes` | `mode: "Off"` refused | `mode` sets it | the agent is always tracked (proposal 12) |
| `read`, `outline`, `find`, `extract_text` | the text extraction specification's views and options (section 2.13) | `view`: `accepted` or `markup` | the text model shared with docx4j and the Python server (ED-008) |
| `comment` | `resolve` | — | resolving a thread is a comment operation here |
| `open`, `list_documents` | present | — | editor only (already in the vocabulary) |

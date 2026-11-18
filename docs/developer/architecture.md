# Architecture

The editor holds a `.docx` in three layers, and each is the source of truth for one thing.

| Layer | What it is | Source of truth for |
|---|---|---|
| Document | [`@docx4j/core-ts`](https://github.com/plutext/docx4j-core-ts)'s `WordprocessingMLPackage`, in the browser or in Node | the file |
| Edit | a ProseMirror document held in a Yjs `Y.XmlFragment` through `y-prosemirror` | the edit session: undo, co-editing, carets |
| View | ProseMirror's `EditorView` on `contenteditable`, its look from CSS generated from the document's styles | nothing: it draws the edit layer |

- **The document layer** unmarshals only the parts the editor shows: the main document part, styles,
  numbering, fonts, theme, settings, headers and footers. The notes and comments parts wait until
  they are first displayed. Every other part stays as bytes and is copied back on save.
- **The edit layer** is a Yjs document from the first keystroke, even with one person editing. Undo
  is a `Y.UndoManager` scoped to the local origin, not ProseMirror's history, so it undoes only your
  own edits in a shared room. Notes, headers and comments are fragments of their own in `Y.Map`s
  keyed by id.
- **The view layer** draws what the editor cannot edit (pictures, opaque content) as node views. It
  doesn't paginate: page boundaries come from the `w:lastRenderedPageBreak` markers Word wrote, and
  printing uses the browser's pages (`docs/user/printing.md`).

`packages/editor-model` is the document and edit layers, headless (it runs in Node with no DOM).
`packages/editor` is the view and the application. `packages/editor-relay` is the co-editing relay,
and `packages/editor-mcp` the agent's companion, which joins a room as a peer.

## The projection

The edit layer holds the part of WordprocessingML the editor can edit: the projection. Fidelity
depends on a few rules.

- **An editable node remembers its element.** It carries the `sourceId` of the element it came from,
  and only the properties the editor can change: direct formatting and references, never formatting
  resolved through styles. Resolved formatting is computed for display only.
- **Export reuses what it can.**
  - An unchanged node writes its original element.
  - A changed node is rebuilt from a deep copy of the original, so properties the editor never
    showed survive.
  - A new node is built with the engine's builders.
- **Nothing is dropped silently.** What the editor cannot edit (fields, math, unknown elements and
  so on) becomes an opaque atom. It is drawn, written back as it was, and listed under Diagnostics.
- **Tracked changes are content.** Insertions and deletions are marks on the text, and a deletion
  keeps its text. No revision id is stored: a revision the export writes anew takes its id from the
  engine.
- **Text is measured in UTF-16 code units**, and a run is never split inside a surrogate pair or a
  grapheme cluster. Direction comes from `w:bidi` and `w:rtl` alone, drawn with CSS logical
  properties.
- **Commands speak the content API, not ProseMirror positions.** Each command has the shape of an
  Office JS call (`Word.*`) and is registered once. The toolbar, keys, palette, console and agent
  tools all bind to the same registry. A test runs each command's Office JS form (its `api` line)
  against the command, and their exports must agree.

## The round trip

A document opened and saved without edits comes back identical part by part. The exceptions are
`[Content_Types].xml` and the `.rels` parts, which the engine regenerates. Before every save, the
editor exports, reloads what it wrote, projects it again and compares; a mismatch blocks the save.

A change to import or export needs a round-trip test over the fixtures in
`packages/editor-model/test/fixtures/` (its README lists them and their sources).

## Speed

Nothing may run over the whole document on every keystroke. A plugin that walks the document either
memoises per top-level block or checks first whether the transaction touched what it cares about.
The benchmarks (`packages/editor-model/test/bench.test.mts`, `packages/editor/e2e/benchmark.spec.ts`)
report figures over a 255-page document; they never fail the suite.

## Where else to look

- `docs/developer/word-divergences.md`: where the editor knowingly differs from Word.
- `docs/developer/upstream-limitations.md`: third-party limitations, with `// UPSTREAM-n:` at each workaround.
- `docs/developer/agent.md`: the agent's tools.
- `docs/developer/localisation.md`: the interface's languages.

Comments in the code cite the project's design records by number and section (for example
"ED-005 section 12.75"). The records themselves are not published.

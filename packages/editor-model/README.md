# @docx4j/editor-model

The headless editable projection of a docx over [`@docx4j/core-ts`](https://github.com/plutext/docx4j-core-ts):
a ProseMirror schema for the editable subset of WordprocessingML, import from the typed tree,
export back to it with every untouched element re-emitted as it was, and a Yjs document as the
edit session's store. Runs in Node and the browser; no DOM. The editor application is
`@docx4j/editor`.

Design: [`docs/developer/architecture.md`](https://github.com/plutext/docx4j-editor/blob/main/docs/developer/architecture.md) in the repository.

```ts
import { EditorDocument } from '@docx4j/editor-model';

const session = await EditorDocument.open(bytes);      // a .docx as Uint8Array
session.doc;                                            // the ProseMirror document (read from the Yjs store)
session.diagnostics;                                    // what is shown read-only, and why
session.replace(editedDoc);                             // one local transaction (a view uses ySyncPlugin instead)
const { bytes: saved, changed } = await session.save(); // export, self-check, zip
```

What the round trip guarantees, and tests: a document opened and saved without edits is
byte-identical part by part, the main document part included; after an edit, every element the
edit did not touch is the original element, and a changed one is rebuilt from a copy of the
original with only the editor's fields written back.

Licence: `GPL-3.0-or-later` (see `LICENSE`); contributions under Apache-2.0 (`CONTRIBUTING.md`).

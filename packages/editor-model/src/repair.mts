// SPDX-License-Identifier: GPL-3.0-or-later
//
// The repairs a save makes to the document it writes (ED-004 sections 3.5 and 6.1). Two peers'
// structural edits, each valid, can merge into a table whose rows have different widths (two
// columns added at once, a row deleted under a merge); neither peer's fixing plugin ran, since
// each saw a valid table, and the view repairs it on the next local change that touches it. A
// save does not wait for that: it writes such a table repaired with `prosemirror-tables`'
// `fixTables`, and the row groups normalised, both deterministic, so every peer holding the same
// state writes the same repair. The store is not changed.
//
// Only what an edit broke is repaired. A table is left as it is when it is the table the session
// imported (an untouched table round-trips byte for byte), when its only problems are column
// widths (`colwidth mismatch`, which a reload shows the same way), and when the table it was
// imported as was already irregular: Word writes rows with fewer cells than the grid on purpose
// (`w:gridBefore`, `w:gridAfter`), and filling them in would change the document.
import { EditorState } from 'prosemirror-state';
import { fixTables, TableMap } from 'prosemirror-tables';
import { Fragment, type Node as PMNode } from 'prosemirror-model';
import { normalizeRowGroups, endCellsWithParagraph } from './commands/tables.mjs';

/** A table's problems other than column widths: rows of different widths, spans that collide or overrun. */
function structuralProblems(table: PMNode): number {
  return (TableMap.get(table).problems ?? []).filter((p) => p.type !== 'colwidth mismatch').length;
}

/** The top-level blocks equal to the node they were imported as, by index. */
function unchangedBlocks(doc: PMNode, imported: (sourceId: string) => PMNode | undefined): Set<number> {
  const out = new Set<number>();
  doc.forEach((block, _pos, i) => {
    const id = block.attrs['sourceId'] as string | null;
    const before = id ? imported(id) : undefined;
    if (before && before.eq(block)) out.add(i);
  });
  return out;
}

/**
 * The document as a save writes it. `imported` gives the node a block was imported as, by
 * `sourceId` (the session's baseline), so a table the edits did not break is never touched.
 */
export function repairedForSave(doc: PMNode, imported: (sourceId: string) => PMNode | undefined): PMNode {
  const blocks: PMNode[] = [];
  let changed = false;
  doc.forEach((block) => {
    if (block.type.name !== 'table' || structuralProblems(block) === 0) { blocks.push(block); return; }
    const id = block.attrs['sourceId'] as string | null;
    const before = id ? imported(id) : undefined;
    if (before && (before.eq(block) || structuralProblems(before) > 0)) { blocks.push(block); return; }
    const alone = doc.type.create(doc.attrs, [block]);
    const fixed = fixTables(EditorState.create({ doc: alone }));
    blocks.push(fixed?.docChanged ? fixed.doc.firstChild! : block);
    changed ||= fixed?.docChanged === true;
  });
  let next = changed ? doc.copy(Fragment.from(blocks)) : doc;
  const grouped = normalizeRowGroups(EditorState.create({ doc: next }).tr);
  if (grouped?.docChanged) next = grouped.doc;
  // A cell ending in a table (ED-005 section 12.109), in a block an edit changed: the paragraph Word's repair
  // appends. A block equal to the one imported is left as it is, so an untouched file round-trips byte for byte.
  const ended = endCellsWithParagraph(EditorState.create({ doc: next }).tr, true, unchangedBlocks(next, imported));
  if (ended?.docChanged) next = ended.doc;
  return next;
}

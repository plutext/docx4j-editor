// SPDX-License-Identifier: GPL-3.0-or-later
//
// A content control's own properties honoured on a local edit (ED-005 section 12.60, `controlrules.mts`): its lock
// (cannot delete, cannot edit contents), its placeholder (typing replaces it) and remove when edited (`w:temporary`).
import { describe, expect, test } from 'vitest';
import type { Node as PMNode } from 'prosemirror-model';
import { EditorDocument, runCommand, readControl, sdtBinding, CellSelection, TableMap, PLACEHOLDER_TEXT, PLACEHOLDER_STYLE, type SdtGroup } from '../src/index.mjs';
import { fixture } from './helpers.mjs';

/** HelloWordOnline.docx's first paragraph: a bookmark's two markers, then "Hello Word..." from `text`. */
async function hello(): Promise<{ session: EditorDocument; text: number }> {
  const session = await EditorDocument.open(await fixture('HelloWordOnline.docx'));
  return { session, text: 1 + 2 };
}

/** The run controls in the document, with their positions. */
function controls(doc: PMNode): { node: PMNode; pos: number }[] {
  const out: { node: PMNode; pos: number }[] = [];
  doc.descendants((n, pos) => { if (n.type.name === 'sdtInline' || n.type.name === 'sdtBlock') out.push({ node: n, pos }); return true; });
  return out;
}

/** A run control around "Hello" with the given properties. */
async function locked(properties: Record<string, unknown>): Promise<{ session: EditorDocument; start: number; inside: number }> {
  const { session, text } = await hello();
  expect(runCommand(session, 'contentControl.insert', { kind: 'RichText' }, { from: text, to: text + 5 }).ok).toBe(true);
  expect(runCommand(session, 'contentControl.properties', properties, { from: text + 2 }).ok).toBe(true);
  return { session, start: text, inside: text + 3 };
}

describe('a control that cannot be edited', () => {
  test('typing, deleting and formatting inside it are refused; outside it, and deleting it whole, are not; tracked too', async () => {
    const { session, start, inside } = await locked({ cannotEdit: true });
    const before = session.doc;
    expect(runCommand(session, 'insertText', { text: 'x' }, { from: inside })).toMatchObject({ ok: false, refused: 'cannotEdit' });
    expect(runCommand(session, 'insertText', { text: '' }, { from: inside, to: inside + 2 })).toMatchObject({ ok: false, refused: 'cannotEdit' });
    expect(runCommand(session, 'font.bold', { value: true }, { from: start + 1, to: start + 6 })).toMatchObject({ ok: false, refused: 'cannotEdit' });
    expect(session.doc.eq(before)).toBe(true);
    // With tracking on, a deletion inside it is a change of its content too.
    session.trackingMode = 'TrackAll';
    expect(runCommand(session, 'insertText', { text: '' }, { from: inside, to: inside + 2 })).toMatchObject({ ok: false, refused: 'cannotEdit' });
    session.trackingMode = 'Off';
    // Just before it and just after it: not inside.
    expect(runCommand(session, 'insertText', { text: '<' }, { from: start }).ok).toBe(true);
    const control = controls(session.doc)[0]!;
    expect(runCommand(session, 'insertText', { text: '>' }, { from: control.pos + control.node.nodeSize }).ok).toBe(true);
    expect(controls(session.doc)[0]!.node.textContent).toBe('Hello');
    // Its properties may change (the Control tab), and it may be deleted whole, as in Word.
    expect(runCommand(session, 'contentControl.properties', { title: 'Locked' }, { from: control.pos + 3 }).ok).toBe(true);
    expect(runCommand(session, 'contentControl.delete', undefined, { from: control.pos + 3 }).ok).toBe(true);
    expect(controls(session.doc)).toHaveLength(0);
  });

  test('a bound value is written into it, as Word fills a locked mapped control', async () => {
    const { session, inside } = await locked({ cannotEdit: true });
    session.pkg.customXmlParts.add('<data><name>Ann</name></data>');
    await session.template.load();
    expect(runCommand(session, 'contentControl.bind', { xpath: '/data/name' }, { from: inside }).ok).toBe(true);
    expect(controls(session.doc)[0]!.node.textContent).toBe('Ann');
    expect(sdtBinding(controls(session.doc)[0]!.node.attrs['sdtPr'])?.xpath).toBe('/data/name');
  });
});

describe('a control that cannot be deleted', () => {
  test('deleting it, or a range around it, is refused; editing inside it is not; unlocked, it goes', async () => {
    const { session, start, inside } = await locked({ cannotDelete: true });
    expect(runCommand(session, 'contentControl.delete', undefined, { from: inside })).toMatchObject({ ok: false, refused: 'cannotDelete' });
    expect(runCommand(session, 'contentControl.delete', { keepContent: true }, { from: inside })).toMatchObject({ ok: false, refused: 'cannotDelete' });
    const end = controls(session.doc)[0]!.pos + controls(session.doc)[0]!.node.nodeSize;
    expect(runCommand(session, 'insertText', { text: '' }, { from: start - 1, to: end + 2 })).toMatchObject({ ok: false, refused: 'cannotDelete' });
    expect(runCommand(session, 'insertText', { text: 'J' }, { from: start + 1, to: start + 2 }).ok).toBe(true);
    expect(controls(session.doc)[0]!.node.textContent).toBe('Jello');
    // Both locks: neither.
    expect(runCommand(session, 'contentControl.properties', { cannotEdit: true }, { from: inside }).ok).toBe(true);
    expect(readControl(controls(session.doc)[0]!.node.attrs['sdtPr'])).toMatchObject({ cannotDelete: true, cannotEdit: true });
    expect(runCommand(session, 'insertText', { text: 'x' }, { from: inside })).toMatchObject({ ok: false, refused: 'cannotEdit' });
    expect(runCommand(session, 'contentControl.delete', undefined, { from: inside })).toMatchObject({ ok: false, refused: 'cannotDelete' });
    // The Control tab unlocks it; then it goes.
    expect(runCommand(session, 'contentControl.properties', { cannotDelete: false, cannotEdit: false }, { from: inside }).ok).toBe(true);
    expect(runCommand(session, 'contentControl.delete', undefined, { from: inside }).ok).toBe(true);
  });

  test('a row group that cannot be deleted keeps its rows, tracking off and on; one row of it may go; one whose contents are locked keeps its cells', async () => {
    const session = await EditorDocument.open(await fixture('tables.docx'));
    // Table 7 has five rows: the group is its first two, so deleting them leaves the table.
    const table = 7;
    let tablePos = 0;
    for (let k = 0; k < table; k++) tablePos += session.doc.child(k).nodeSize;
    const rowsOfGroup = (): CellSelection => {
      const map = TableMap.get(session.doc.child(table));
      return CellSelection.create(session.doc, tablePos + 1 + map.map[0]!, tablePos + 1 + map.map[map.width]!);
    };
    const a = tablePos + 1 + TableMap.get(session.doc.child(table)).map[0]!;
    expect(runCommand(session, 'contentControl.wrapRows', undefined, rowsOfGroup()).ok).toBe(true);
    expect(runCommand(session, 'contentControl.properties', { cannotDelete: true }, rowsOfGroup()).ok).toBe(true);
    const inCell = a + 2;
    const rows = (): number => session.doc.child(table).childCount;
    const before = rows();
    // All the group's rows: the control would go.
    expect(runCommand(session, 'table.deleteRows', undefined, rowsOfGroup())).toMatchObject({ ok: false, refused: 'cannotDelete' });
    session.trackingMode = 'TrackAll';
    expect(runCommand(session, 'table.deleteRows', undefined, rowsOfGroup())).toMatchObject({ ok: false, refused: 'cannotDelete' });
    session.trackingMode = 'Off';
    expect(rows()).toBe(before);
    expect(runCommand(session, 'insertText', { text: 'x' }, { from: inCell }).ok).toBe(true);
    // Its contents locked: no cell edited, no row deleted.
    expect(runCommand(session, 'contentControl.properties', { cannotDelete: false, cannotEdit: true }, rowsOfGroup()).ok).toBe(true);
    expect(runCommand(session, 'insertText', { text: 'y' }, { from: inCell })).toMatchObject({ ok: false, refused: 'cannotEdit' });
    expect(runCommand(session, 'table.deleteRows', undefined, { from: inCell })).toMatchObject({ ok: false, refused: 'cannotEdit' });
    // Only the control locked: one of its rows is content, and goes; the control stays on the other.
    expect(runCommand(session, 'contentControl.properties', { cannotDelete: true, cannotEdit: false }, rowsOfGroup()).ok).toBe(true);
    expect(runCommand(session, 'table.deleteRows', undefined, { from: inCell }).ok).toBe(true);
    expect(rows()).toBe(before - 1);
    expect(((session.doc.child(table).child(0).attrs['sdt'] as SdtGroup[] | null) ?? []).length).toBe(1);
  });
});

describe('the placeholder and remove when edited', () => {
  test('typing in a control showing its placeholder replaces it, untracked, in one undo step', async () => {
    const { session, text } = await hello();
    expect(runCommand(session, 'contentControl.insert', { kind: 'RichText' }, { from: text }).ok).toBe(true);
    const made = (): PMNode => controls(session.doc)[0]!.node;
    expect(made().textContent).toBe(PLACEHOLDER_TEXT);
    session.undoManager.stopCapturing();
    expect(runCommand(session, 'insertText', { text: 'Hi' }, { from: text + 1 + 5 }).ok).toBe(true);
    expect(made().textContent).toBe('Hi');
    expect(readControl(made().attrs['sdtPr']).showingPlaceholder).toBe(false);
    expect(made().child(0).marks.some((m) => m.type.name === 'runStyle' && m.attrs['id'] === PLACEHOLDER_STYLE)).toBe(false);
    session.undoManager.undo();
    expect(made().textContent).toBe(PLACEHOLDER_TEXT);
    expect(readControl(made().attrs['sdtPr']).showingPlaceholder).toBe(true);
    // With tracking on, what is typed is an insertion; the placeholder goes for real, being no content.
    session.trackingMode = 'TrackAll';
    session.trackingAuthor = 'Reviewer';
    expect(runCommand(session, 'insertText', { text: 'Yo' }, { from: text + 1 }).ok).toBe(true);
    expect(made().textContent).toBe('Yo');
    expect(made().child(0).marks.map((m) => m.type.name)).toContain('insertion');
  });

  test('a bound control showing its placeholder: what is typed replaces it and is written to the node', async () => {
    const { session, text } = await hello();
    const part = session.pkg.customXmlParts.add('<data><empty/></data>');
    await session.template.load();
    expect(runCommand(session, 'contentControl.insertBound', { xpath: '/data/empty' }, { from: text }).ok).toBe(true);
    expect(controls(session.doc)[0]!.node.textContent).toBe(PLACEHOLDER_TEXT);
    expect(runCommand(session, 'insertText', { text: 'Zed' }, { from: text + 1 }).ok).toBe(true);
    expect(controls(session.doc)[0]!.node.textContent).toBe('Zed');
    expect(part.selectSingleNode('/data/empty')!.text).toBe('Zed');
  });

  test('remove when edited: the first edit inside removes the control and keeps its content, in one undo step', async () => {
    const { session, inside } = await locked({ removeWhenEdited: true });
    const text = session.doc.child(0).textContent;
    session.undoManager.stopCapturing();
    expect(runCommand(session, 'insertText', { text: 'x' }, { from: inside }).ok).toBe(true);
    expect(controls(session.doc)).toHaveLength(0);
    expect(session.doc.child(0).textContent).toBe(text.slice(0, 2) + 'x' + text.slice(2));   // "He|llo"
    session.undoManager.undo();
    expect(controls(session.doc)).toHaveLength(1);
    expect(session.doc.child(0).textContent).toBe(text);
    // Its properties changed are no edit of its content: it stays.
    expect(runCommand(session, 'contentControl.properties', { title: 'Temp' }, { from: inside }).ok).toBe(true);
    expect(controls(session.doc)).toHaveLength(1);
    // With its placeholder showing: what is typed stays, without the control.
    const fresh = await hello();
    expect(runCommand(fresh.session, 'contentControl.insert', { kind: 'RichText' }, { from: fresh.text }).ok).toBe(true);
    expect(runCommand(fresh.session, 'contentControl.properties', { removeWhenEdited: true }, { from: fresh.text + 2 }).ok).toBe(true);
    expect(runCommand(fresh.session, 'insertText', { text: 'Hi' }, { from: fresh.text + 1 }).ok).toBe(true);
    expect(controls(fresh.session.doc)).toHaveLength(0);
    expect(fresh.session.doc.child(0).textContent.slice(0, 2)).toBe('Hi');
  });
});

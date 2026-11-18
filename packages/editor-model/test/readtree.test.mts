// SPDX-License-Identifier: GPL-3.0-or-later
//
// ED-005 section 8.9, E4.f step F3: `readTree`, the tree as the document is now for the agent's
// reading tools and its dry runs. It reads the edit layer's current state; with `copy`, a change
// made in the callback (tracked, when asked) leaves the document, the store and the saved file as
// they were.
import { describe, test, expect } from 'vitest';
import { PlainText, Texts, WordprocessingMLPackage } from '@docx4j/core-ts';
import { EditorDocument, runCommand } from '../src/index.mjs';
import { fixture, comparePackages } from './helpers.mjs';

describe('readTree (ED-005 section 8.9)', () => {
  test('reads the document as it is now, edits in the edit layer included', async () => {
    const session = await EditorDocument.open(await fixture('tracked-changes.docx'));
    let at = -1;
    session.doc.forEach((n, pos) => { if (at < 0 && n.textContent === 'Some text.') at = pos + 1; });
    runCommand(session, 'insertText', { text: 'Edited: ' }, { from: at });
    const outline = await session.readTree(async ({ body }) => (await body.outline()).paragraphs.map((p) => p.text));
    expect(outline).toContain('Edited: Some text.');
  });

  test('with copy, a tracked change made in the callback is reported and leaves the document as it was', async () => {
    const session = await EditorDocument.open(await fixture('tracked-changes.docx'));
    const before = session.doc;
    const changes = await session.readTree(({ body }) => {
      body.insertParagraph('A dry run.', 'End');
      return body.getTrackedChanges().filter((c) => c.author === 'Claude').map((c) => c.text);
    }, { copy: true, tracking: { author: 'Claude', mode: 'TrackAll' } });
    expect(changes.join('')).toContain('A dry run.');
    expect(session.doc.eq(before)).toBe(true);
    const again = await session.readTree(async ({ body }) => (await body.outline()).paragraphs.some((p) => p.text === 'A dry run.'));
    expect(again).toBe(false);
    expect((await session.save()).changed).toBe(false);
  });
});

describe('the tree bindings\' doc and elementOf (ED-005 section 12.27): a node of the edit layer to its address in the tree', () => {
  for (const name of ['tracked-changes.docx', 'tables.docx', 'nested-table.docx']) {
    test(`${name}: every paragraph node resolves to the address the engine's outline gives it, before and after an edit`, async () => {
      const session = await EditorDocument.open(await fixture(name));
      const check = async (): Promise<number> => session.readTree(async ({ doc, elementOf, body }) => {
        const expected = (await body.outline()).paragraphs.map((p) => p.address);
        const got: string[] = [];
        doc.descendants((node) => {
          if (node.type.name !== 'paragraph') return true;
          const element = elementOf(node);
          got.push(element ? body.addressOf({ element, container: [] }) ?? 'none' : 'unmapped');
          return false;
        });
        // Every node is one of the outline's paragraphs, in document order; the outline also has the
        // paragraphs of a vertical merge's continuation cells, which the edit layer keeps as side data.
        let last = -1;
        for (const address of got) {
          const i = expected.indexOf(address);
          expect(i, address).toBeGreaterThan(last);
          last = i;
        }
        return got.length;
      });
      expect(await check()).toBeGreaterThan(3);
      // An edit in the edit layer: the export rebuilds that block, and the bridge follows it.
      let at = -1;
      session.doc.descendants((n, pos) => { if (at < 0 && n.type.name === 'paragraph' && n.textContent.length > 0) at = pos + 1; return at < 0; });
      runCommand(session, 'insertText', { text: 'Edited ' }, { from: at });
      await check();
    });
  }
});

describe('the text model\'s package (ED-008 section 3.1): the engine\'s Texts reads the session\'s tree, and the parts are left as they were', () => {
  test('an edit not yet handed to the part is read; the file is not; a save afterwards is byte for byte where it was', async () => {
    const original = await fixture('tracked-changes.docx');
    const session = await EditorDocument.open(original);
    // Before any edit the main part was never unmarshalled: the view reads the session's tree all the same.
    const unedited = await session.readTree(async ({ textPackage }) => PlainText.of(await Texts.of(textPackage)));
    expect(unedited).toBe(PlainText.of(await Texts.of(await WordprocessingMLPackage.load(original))));
    expect(comparePackages(original, (await session.save()).bytes)).toEqual([]);
    let at = -1;
    session.doc.forEach((n, pos) => { if (at < 0 && n.textContent === 'Some text.') at = pos + 1; });
    runCommand(session, 'insertText', { text: 'Edited: ' }, { from: at });
    const read = await session.readTree(async ({ textPackage, pkg }) => ({ view: PlainText.of(await Texts.of(textPackage)), part: PlainText.of(await Texts.of(pkg)) }));
    expect(read.view).toContain('Edited: Some text.');
    // the package itself still reads the file: the part was never handed the tree
    expect(read.part).not.toContain('Edited:');
    expect(session.main.isUnmarshalled).toBe(false);
  });

  test('after a tree run the main part holds the tree, and the view is the package itself', async () => {
    const session = await EditorDocument.open(await fixture('tracked-changes.docx'));
    await session.runOnTree(({ body }) => { body.insertParagraph('From a tree run.', 'End'); });
    const same = await session.readTree(({ textPackage, pkg }) => textPackage === pkg);
    expect(same).toBe(true);
    expect(await session.readTree(async ({ textPackage }) => PlainText.of(await Texts.of(textPackage)))).toContain('From a tree run.');
  });

  test('over a copy, the copy\'s body: a dry run\'s change is read, and the document is not changed', async () => {
    const session = await EditorDocument.open(await fixture('tracked-changes.docx'));
    const text = await session.readTree(async ({ body, textPackage }) => {
      body.insertParagraph('A dry run.', 'End');
      return PlainText.of(await Texts.of(textPackage));
    }, { copy: true });
    expect(text).toContain('A dry run.');
    expect(await session.readTree(async ({ textPackage }) => PlainText.of(await Texts.of(textPackage)))).not.toContain('A dry run.');
  });

  test('with stories, a note and a header edited in the session are read; without, the parts as they were', async () => {
    for (const [name, kind] of [['footnote.docx', 'footnote'], ['headers.docx', 'header']] as const) {
      const original = await fixture(name);
      const session = await EditorDocument.open(original);
      const story = kind === 'footnote' ? (await session.stories.footnotes()).find((n) => !n.orphan)! : (await session.stories.headers()).find((h) => h.kind === 'header' && h.doc.textContent.length > 0)!;
      runCommand(session, 'insertText', { text: 'EditedStory ' }, { from: kind === 'footnote' ? startOfNoteText(story.doc) : 1 }, story);
      const options = { include: ['headers'] } as const;
      const before = await session.readTree(async ({ textPackage }) => PlainText.of(await Texts.of(textPackage, options)));
      expect(before, name).not.toContain('EditedStory');
      const after = await session.readTree(async ({ textPackage }) => PlainText.of(await Texts.of(textPackage, options)), { stories: true });
      expect(after, name).toContain('EditedStory');
      // the part written is one the save writes, as it would have been; the self-check holds
      const saved = await session.save();
      expect(saved.changed, name).toBe(true);
      expect(saved.selfCheck?.equal, name).toBe(true);
      const reopened = await EditorDocument.open(saved.bytes);
      expect(await reopened.readTree(async ({ textPackage }) => PlainText.of(await Texts.of(textPackage, options)), { stories: true }), name).toContain('EditedStory');
    }
  });

  test('with stories over an unedited document, nothing is written: the save is byte for byte', async () => {
    for (const name of ['footnote.docx', 'headers.docx', 'comments-two.docx']) {
      const original = await fixture(name);
      const session = await EditorDocument.open(original);
      await session.stories.comments();
      await session.readTree(async ({ textPackage }) => PlainText.of(await Texts.of(textPackage, { include: ['headers', 'footers'], comments: true })), { stories: true });
      const saved = await session.save();
      expect(saved.changed, name).toBe(false);
      expect(comparePackages(original, saved.bytes), name).toEqual([]);
    }
  });
});

/** The position just inside a note's first paragraph's text, after its own note mark. */
function startOfNoteText(doc: import('prosemirror-model').Node): number {
  let pos = 1;
  doc.firstChild!.forEach((child, offset) => { if (pos === 1 && child.isText) pos = 1 + offset; });
  return pos;
}

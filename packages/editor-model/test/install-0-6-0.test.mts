// SPDX-License-Identifier: GPL-3.0-or-later
//
// The checks ED-008 section 8.8 asked for at the core-ts 0.6.0 install, each over the engine's new
// writes through a tree run (the agent's path): item 10, a tracked deletion taking an equation in
// Word's form, listed among the entries shown in Word only and emptied by Accept All; D7, every write
// into a bound control writing its custom XML node, the controls bound to it then agreeing with the
// node (nothing to refresh, nothing written twice); D3, a Replace over a whole control showing its
// placeholder taking the control, tracked, and Accept All removing it.
import { describe, test, expect } from 'vitest';
import { EditorState } from 'prosemirror-state';
import { EditorDocument, runCommandAsync, preservedChangesOf, refreshBoundControls } from '../src/index.mjs';
import { fixture } from './helpers.mjs';
import { mainXmlOf } from './api-cases.mjs';

const TRACKING = { author: 'Claude', mode: 'TrackAll' as const, date: new Date('2026-10-26T10:00:00Z') };

/** The entries shown in Word only, as the review panel lists them. */
const preserved = (session: EditorDocument) => preservedChangesOf(session.doc, (id) => session.ctx.ids.get(id)?.value);

describe('core-ts 0.6.0 install checks (ED-008 section 8.8)', () => {
  test('item 10: a tracked deletion of a paragraph holding an equation is Word\'s form, a Word-only entry; Accept All leaves no emptied equation', async () => {
    const session = await EditorDocument.open(await fixture('loadAndSave.docx'));
    let index = -1;
    session.doc.forEach((block, _pos, i) => {
      if (index >= 0 || block.type.name !== 'paragraph') return;
      block.descendants((n) => { if (n.type.name === 'opaqueInline' && (n.attrs['kind'] === 'oMath' || n.attrs['kind'] === 'oMathPara')) index = i; return index < 0; });
    });
    expect(index).toBeGreaterThanOrEqual(0);
    const before = (await mainXmlOf(session)).match(/<m:oMath>/g)?.length ?? 0;
    expect(before).toBeGreaterThan(0);
    const run = await session.runOnTree(({ body }) => { body.paragraphAt(`body/${index}`)!.delete(); }, undefined, TRACKING);
    expect(run.refused).toBeUndefined();
    const xml = await mainXmlOf(session);
    // Word's form (core-ts check 57): the deletion inside each math run, not one w:del around the equation.
    expect(xml).toMatch(/<m:ctrlPr><w:del |<m:r><w:del /);
    expect(preserved(session).map((p) => p.kind)).toEqual(['mathDel']);
    expect((await runCommandAsync(session, 'trackedChanges.acceptAll')).ok).toBe(true);
    const after = await mainXmlOf(session);
    expect(after).not.toMatch(/<m:oMath\s*\/>|<m:oMath><\/m:oMath>/);
    expect(after.match(/<m:oMath>/g)?.length ?? 0).toBeLessThan(before);
    expect(after).not.toContain('<w:del ');
  });

  test('D7: a write into a bound control through a tree run writes its node, and the controls bound to it then need no refresh', async () => {
    const session = await EditorDocument.open(await fixture('opendope/binding-simple.docx'));
    await session.template.load();
    const run = await session.runOnTree(({ body }) => {
      const control = body.contentControls.find((c) => c.xmlMapping?.customXmlNode !== undefined);
      expect(control).toBeDefined();
      control!.insertText('Perth', 'Replace');
      return control!.xmlMapping!.customXmlNode!.text;
    });
    expect(run.refused).toBeUndefined();
    expect(run.value).toBe('Perth');
    expect(session.doc.textContent).toContain('Perth');
    await session.template.load();
    const state = EditorState.create({ doc: session.doc });
    expect(refreshBoundControls(session, state)).toBeUndefined();
  });

  test('D3: a tracked Replace over a whole control showing its placeholder takes the control, a Word-only entry; Accept All removes it', async () => {
    const session = await EditorDocument.open(await fixture('tracked-changes.docx'));
    // A run-level control at the end of "Some text.", showing its placeholder, made as the agent's insert_content_control makes one.
    const made = await session.runOnTree(({ body }) => {
      const p = body.paragraphs.find((q) => q.text === 'Some text.')!;
      const control = p.getRange('End').insertContentControl();
      return { id: control.id, showing: control.isShowingPlaceholder, address: body.addressOf(p)! };
    });
    expect(made.refused).toBeUndefined();
    expect(made.value!.showing).toBe(true);
    // The paragraph's whole content replaced, the control inside it taken whole (CR-002 section 63.12 D3).
    const replaced = await session.runOnTree(({ body }) => {
      body.paragraphAt(made.value!.address)!.getRange('Content').insertText('Replaced', 'Replace');
    }, undefined, TRACKING);
    expect(replaced.refused).toBeUndefined();
    const xml = await mainXmlOf(session);
    expect(xml).toContain('<w:customXmlDelRangeStart');
    expect(preserved(session).map((p) => p.kind)).toEqual(['holderDel']);
    expect((await runCommandAsync(session, 'trackedChanges.acceptAll')).ok).toBe(true);
    const after = await mainXmlOf(session);
    expect(after).not.toContain(`<w:id w:val="${made.value!.id}"/>`);
    expect(after).not.toContain('customXmlDelRange');
    expect(after).toContain('Replaced');
  });
});

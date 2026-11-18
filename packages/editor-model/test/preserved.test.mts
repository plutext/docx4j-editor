// SPDX-License-Identifier: GPL-3.0-or-later
//
// ED-005 section 12.26: the revisions shown in Word only, resolved through the engine since core-ts
// 0.3.0 (CR-002 section 35). A review panel entry names its revision by its top-level block, kind
// and place (`preservedTargetOf`); `trackedChange.accept` and `.reject` with that target resolve it
// in their settle half, a tree run. Held here to the engine resolving the same change picked by its
// own listing's order, over Word's files; Accept all and Reject all are `agreement.test.mts`'s.
import { describe, test, expect } from 'vitest';
import { EditorDocument, runCommand, runCommandAsync, preservedChangesOf, preservedGroupsOf, preservedTargetOf, type PreservedGroup } from '../src/index.mjs';
import { fixture } from './helpers.mjs';
import { mainXmlOf, runLine, diff } from './api-cases.mjs';
import { room } from './room-harness.mjs';

const FILES = ['revisions/revisions-word2010.docx', 'revisions/revisions-word15.docx', 'revisions/check28/28a-changes.docx'];

function groupsOf(session: EditorDocument): PreservedGroup[] {
  return preservedGroupsOf(session.doc, preservedChangesOf(session.doc, (id) => session.ctx.ids.get(id)?.value));
}

/** The engine's pick of the same change: its listing's changes of the kind, in its order, the `index`th. */
const ENGINE_KIND: Record<string, string> = {
  tableProperties: "c.target.kind === 'tableProperties'",
  cellIns: "c.target.kind === 'cell' && c.target.cell === 'ins'",
  cellDel: "c.target.kind === 'cell' && c.target.cell === 'del'",
  sectPrChange: "c.target.kind === 'sectionProperties'",
  markRPrChange: "c.target.kind === 'markProperties'",
};

const withoutIds = (xml: string): string => xml.replace(/ w:id="\d+"/g, '').replace(/<w:proofErr [^>]*\/>/g, '');

describe('an entry shown in Word only, resolved through the engine (ED-005 section 12.26)', () => {
  for (const name of FILES) {
    test(`${name}: each entry accepted and rejected alone, as the engine resolves that change`, async () => {
      const probe = await EditorDocument.open(await fixture(name));
      const groups = groupsOf(probe);
      const resolvable = groups.filter((g) => preservedTargetOf(probe.doc, g, groups) !== undefined);
      expect(resolvable.length).toBeGreaterThan(0);
      for (const group of resolvable) {
        const target = preservedTargetOf(probe.doc, group, groups)!;
        const index = resolvable.filter((g) => g.kind === group.kind && g.pos < group.pos).length;
        for (const action of ['accept', 'reject'] as const) {
          const ours = await EditorDocument.open(await fixture(name));
          ours.trackingMode = 'Off';
          const r = await runCommandAsync(ours, `trackedChange.${action}`, { preserved: target });
          expect(r.ok, `${group.kind} ${action}`).toBe(true);
          // It is gone from the listing, and nothing else is.
          expect(groupsOf(ours).map((g) => g.kind)).toEqual(groups.filter((g) => g !== group).map((g) => g.kind));
          const theirs = await EditorDocument.open(await fixture(name));
          theirs.trackingMode = 'Off';
          const pieces = 'body.getTrackedChanges().flatMap((c) => (c.target.kind === \'group\' ? c.target.pieces : [c]))';
          await runLine(theirs, `${pieces}.filter((c) => ${ENGINE_KIND[group.kind]})[${index}].${action}();`, { from: 1 });
          const [a, b] = [withoutIds(await mainXmlOf(ours)), withoutIds(await mainXmlOf(theirs))];
          expect(a === b, `${group.kind} ${action}:\n  ours:   ${diff(a, b)[0]}\n  engine: ${diff(a, b)[1]}`).toBe(true);
        }
      }
    });
  }

  test("an author's all takes their entries shown in Word only; everyone's all, the rest", async () => {
    const session = await EditorDocument.open(await fixture('revisions/revisions-word15.docx'));
    session.trackingMode = 'Off';
    expect(groupsOf(session).map((g) => g.author)).toEqual(['Author A', 'Author A']);
    expect((await runCommandAsync(session, 'trackedChanges.acceptAll', { author: 'Author B' })).ok).toBe(true);
    expect(groupsOf(session).length).toBe(2);
    expect((await runCommandAsync(session, 'trackedChanges.acceptAll', { author: 'Author A' })).ok).toBe(true);
    expect(groupsOf(session)).toEqual([]);
  });

  test('the synchronous runner refuses a command whose tree run has work, and runs one whose has none', async () => {
    const session = await EditorDocument.open(await fixture('revisions/revisions-word15.docx'));
    expect(() => runCommand(session, 'trackedChanges.acceptAll')).toThrow(/runCommandAsync/);
    const plain = await EditorDocument.open(await fixture('tracked-changes.docx'));
    expect(runCommand(plain, 'trackedChanges.acceptAll').ok).toBe(true);
  });

  test('only entries shown in Word only: all is enabled, and resolves them', async () => {
    const session = await EditorDocument.open(await fixture('revisions/revisions-word15.docx'));
    session.trackingMode = 'Off';
    expect((await runCommandAsync(session, 'trackedChanges.rejectAll')).ok).toBe(true);
    expect(groupsOf(session)).toEqual([]);
    // Nothing left: all is not enabled.
    expect((await runCommandAsync(session, 'trackedChanges.acceptAll')).ok).toBe(false);
  });

  test('in a room, an entry one peer accepts is gone on the other', async () => {
    const r = await room(await fixture('revisions/revisions-word15.docx'));
    const groups = groupsOf(r.joiner);
    const table = groups.find((g) => g.kind === 'tableProperties')!;
    expect((await runCommandAsync(r.joiner, 'trackedChange.accept', { preserved: preservedTargetOf(r.joiner.doc, table, groups) })).ok).toBe(true);
    expect(groupsOf(r.creator).map((g) => g.kind)).toEqual(['sectPrChange']);
    expect(groupsOf(r.joiner).map((g) => g.kind)).toEqual(['sectPrChange']);
  });
});

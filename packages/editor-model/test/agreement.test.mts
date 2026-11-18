// SPDX-License-Identifier: GPL-3.0-or-later
//
// ED-005 section 5.8, step C5: the agreement suite. Two trackers write revisions, the editor's
// (typing and commands, `track.mts`) and the engine's (every tree run: the console, macros, the
// ask mode, the agent), and they must write the same markup, or a document edited by both
// carries two dialects into Word. Every case of ED-003 section 7.2 runs with tracking on, the
// same author and a fixed date on both sides, once through the command and once as its `api` line
// through `runOnTree`, and the main parts are compared, revision ids aside (the export makes its
// own, section 5.7). A difference is a finding recorded in section 12 and fixed on whichever side
// is wrong; an engine bug goes to core-ts, and the case is marked `knownTracked` until a release
// carries the fix.
import { describe, test, expect } from 'vitest';
import { WordprocessingMLPackage } from '@docx4j/core-ts';
import { EditorDocument, runCommandAsync, commandNamed, stateFor, trackedChangesOf } from '../src/index.mjs';
import { compareLines, mainXmlOf, runLine, diff } from './api-cases.mjs';
import { fixture } from './helpers.mjs';

const DATE = new Date('2026-09-28T10:00:00Z');

/** Tracking on, as Author C at a fixed minute, for both the command and the line. */
function tracking(session: EditorDocument): void {
  session.trackingMode = 'TrackAll';
  session.trackingAuthor = 'Author C';
  session.trackedChangeDate = DATE;
}

/**
 * Revision ids, which each side's export numbers for itself; and whether the root declares `w16du`,
 * which the engine's tracker does when it writes a date and keeps when the change is dropped again
 * (a link's look taken off, case H), and the editor's export does when it writes one (core-ts 0.3.0,
 * ED-005 section 12.26): a declaration, no content.
 */
const withoutIds = (xml: string): string => replacementOrder(xml.replace(/ w:id="\d+"/g, '')
  .replace(/^(.*?<w:document [^>]*?) xmlns:w16du="[^"]*"/, '$1').replace(/^(.*?<w:document [^>]*? mc:Ignorable="[^"]*?) w16du"/, '$1"'));

/**
 * A tracked replacement's two halves in one order (ED-008 section 8.8 item 9, Jason 2026-10-26: "state the
 * equivalence in the compare"). The editor's recorder writes a person's replacement as Word does when a person
 * types over a selection, the deletion first (checks 24 to 26); the engine's `insertText(…, 'Replace')` writes
 * the insertion first, as Word 365 does for Office JS (check 54 case 4; core-ts 0.6.0, CR-014 phase 3). Both are
 * Word's, Word reads and resolves both the same, so an insertion followed by a deletion of the same author and
 * date is compared as the deletion followed by the insertion.
 */
const replacementOrder = (xml: string): string => xml.replace(
  /<w:ins( [^>]*)>((?:(?!<\/w:ins>).)*?)<\/w:ins><w:del( [^>]*)>((?:(?!<\/w:del>).)*?)<\/w:del>/g,
  (whole, insAttrs: string, ins: string, delAttrs: string, del: string) => (insAttrs === delAttrs ? `<w:del${delAttrs}>${del}</w:del><w:ins${insAttrs}>${ins}</w:ins>` : whole),
);

describe('the editor and the engine track alike (ED-005 section 5.8)', () => {
  test('every api line with tracking on writes the revisions its command writes', async () => {
    const { findings, checked } = await compareLines(await fixture('tracked-changes.docx'), { prepare: tracking, known: (c) => c.knownTracked, normalize: withoutIds });
    expect(checked).toBeGreaterThan(20);
    expect(findings).toEqual([]);
  });
});

/** The Word-made files with revisions pending (Word's own resolutions of them are what `review.test.mts` holds the editor to). */
const PENDING = [
  'tracked-changes.docx', 'revisions/revisions-word2010.docx', 'revisions/revisions-word15.docx',
  ...['01-type-inside-a', '03-type-at-start', '04-type-at-end', '06-delete-part-a', '08-delete-all', '12-partner-missing', '13-moveTo-no-ranges'].map((c) => `revisions/check20/${c}.docx`),
  'revisions/check22/22-deleted.docx', 'revisions/check25/25a-bullets.docx', 'revisions/check25/25e-bullet-demote.docx',
];

/** Word's spelling marks, which the editor's projection does not carry (`import.mts`) and Word's own saves after a resolution do not have. */
const withoutProofing = (xml: string): string => xml.replace(/<w:proofErr [^>]*\/>/g, '');

/**
 * Differences of the engine's over whole classes of file, each normalised away on both sides; each
 * must still be seen, so its entry goes when the engine agrees (ED-005 section 12.7 item 3). The
 * editor's side is Word's (the test after this one).
 */
const KNOWN_CLASSES: { name: string; normalize: (xml: string) => string; reason: string }[] = [
  // None since core-ts 0.3.0: the joined paragraph's id, an empty w:trPr and a rejected list paragraph agree (ED-005 section 12.26).
];

/** Differences known for one file, keyed `file action`, that the classes do not cover. */
const KNOWN_RESOLUTION: Record<string, string> = {
  // None since core-ts 0.3.0 (ED-005 section 12.26).
};

async function resolved(bytes: Uint8Array, action: 'accept' | 'reject', by: 'command' | 'engine'): Promise<string> {
  const session = await EditorDocument.open(bytes);
  if (by === 'command') {
    const command = action === 'accept' ? 'trackedChanges.acceptAll' : 'trackedChanges.rejectAll';
    // The transaction alone resolves everything the edit layer lists, so the settle half's tree run
    // (the engine's "all" over the revisions shown in Word only, ED-005 section 12.26) hides nothing.
    const state = stateFor(session, { from: 1 });
    let after = state;
    commandNamed(command)!.run(undefined)(state, (tr) => { after = state.apply(tr); });
    expect(trackedChangesOf(after.doc), `the edit layer's changes left by ${action} all`).toEqual([]);
    expect((await runCommandAsync(session, command, undefined, { from: 1 })).ok).toBe(true);
  }
  else await runLine(session, action === 'accept' ? 'body.acceptAll();' : 'body.rejectAll();', { from: 1 });
  return withoutProofing(withoutIds(await mainXmlOf(session)));
}

describe('the editor and the engine resolve alike (ED-005 section 5.8)', () => {
  test('accept all and reject all, over every Word file with changes pending: the same main part both ways', async () => {
    const findings: string[] = [];
    const seen = new Set<string>();
    const all = (x: string, except?: string): string => KNOWN_CLASSES.reduce((out, c) => (c.name === except ? out : c.normalize(out)), x);
    for (const name of PENDING) {
      const bytes = await fixture(name);
      for (const action of ['accept', 'reject'] as const) {
        const [raw, rawEngine] = await Promise.all([resolved(bytes, action, 'command'), resolved(bytes, action, 'engine')]);
        const [a, b] = [all(raw), all(rawEngine)];
        if (a === b) for (const c of KNOWN_CLASSES) if (all(raw, c.name) !== all(rawEngine, c.name)) seen.add(c.name);
        const known = KNOWN_RESOLUTION[`${name} ${action}`];
        if (known !== undefined) { if (a === b) findings.push(`${name} ${action}: now agrees with the engine; remove its known difference (${known})`); continue; }
        if (a !== b) findings.push(`${name} ${action}:\n  command: ${diff(a, b)[0]}\n  engine:  ${diff(a, b)[1]}`);
      }
    }
    for (const c of KNOWN_CLASSES) if (!seen.has(c.name)) findings.push(`${c.name}: no longer differs anywhere; remove it (${c.reason})`);
    expect(findings).toEqual([]);
  });

  test('the editor\'s side is Word\'s: after accept all and reject all, the paragraphs keep Word\'s ids and a row no empty trPr (check 18)', async () => {
    const bytes = await fixture('revisions/revisions-word15.docx');
    const idsOf = (xml: string): string[] => [...xml.matchAll(/<w:p [^>]*w14:paraId="(\w+)"/g)].map((m) => m[1]!);
    for (const action of ['accept', 'reject'] as const) {
      const ours = await resolved(bytes, action, 'command');
      const word = await (await WordprocessingMLPackage.load(await fixture(`revisions/check18/${action}-all.docx`))).getMainDocumentPart().getXml();
      expect(idsOf(ours), action).toEqual(idsOf(word));
      expect(ours).not.toContain('<w:trPr/>');
    }
  });
});

// Bound controls whose content is older than their data (ED-005 section 12.89): the counts the application offers a
// refresh for at open, over core-ts check 38's files (`fixtures/pictures/`: the templates stale, the engine's bound
// saves not) and the invoice sample's text controls.
import { describe, expect, test } from 'vitest';
import { EditorDocument, stateFor, staleBoundControls, textOfMarkup, applyBoundRefresh } from '../src/index.mjs';
import { fixture } from './helpers.mjs';

async function counts(name: string): Promise<{ text: number; xhtml: number; picture: number; total: number }> {
  const session = await EditorDocument.open(await fixture(name));
  await session.template.load({ engine: false });
  return staleBoundControls(session, stateFor(session));
}

describe('stale bound controls', () => {
  test("Word's picture control: the template holds a placeholder (stale), the engine's bound save the node's image (not)", async () => {
    expect(await counts('pictures/38a-picture-template.docx')).toMatchObject({ picture: 1, total: 1 });
    expect(await counts('pictures/38a-picture-bound.docx')).toMatchObject({ picture: 0, total: 0 });
  });
  test('the handler controls: the three with a width are stale in the template, none in the bound save; the floating one and the SVG one are left out', async () => {
    expect(await counts('opendope/xpaths-qualified-attributes.docx')).toMatchObject({ picture: 3, total: 3 });
    expect(await counts('pictures/38b-handler-bound.docx')).toMatchObject({ picture: 0, total: 0 });
  });
  test('text controls: the count is what the refresh changes, and nothing after it', async () => {
    const session = await EditorDocument.open(await fixture('invoice.docx'));
    await session.template.load({ engine: false });
    const before = await staleBoundControls(session, stateFor(session));
    const changed = applyBoundRefresh(session, stateFor(session));
    expect(changed).toBe(before.text > 0);
    expect((await staleBoundControls(session, stateFor(session))).text).toBe(0);
  });
  test('the text of markup', () => {
    expect(textOfMarkup('<p>Jane <b>Doe</b></p><ul><li>one</li><li>two &amp; three</li></ul>')).toBe('Jane Doe one two & three');
  });
});

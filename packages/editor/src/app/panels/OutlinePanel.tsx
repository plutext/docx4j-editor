// SPDX-License-Identifier: GPL-3.0-or-later
//
// The outline as an address table (ED-001 section 3.4's addresses; ED-002 section 6.4): what
// E4's tools and the console's `selection.addresses` speak. A click selects the block.
//
// The engine's `outline()` is async since core-ts 0.6.0 and reads the text model (CR-012 D3; ED-008
// sections 8.7 and 8.9): the panel reads it in an effect, memoised per top-level block. It reads the
// tree the engine holds, which the projection came from and which a save or a tree run brings up to
// date with the edits (ED-002 section 8 item 2: an unedited document keeps its source bytes, so the
// edit layer's blocks reach the part only then); the engine reads only blocks that sit in that tree
// (`blockIndexOf`), so a keystroke changes no entry until the next save or run, as before the install.
// At the first paint every block is read under one context; afterwards only the blocks whose element
// is another object are read again, with `{ blocks }` under a fresh context (a context is a snapshot of
// the package), and the whole outline when the block list changed shape or any entry joins others.
import { useEffect, useState } from 'react';
import { Body, Texts, type Outline, type OutlineParagraph, type OutlineTable } from '@docx4j/core-ts';
import type { EditorDocument } from '@docx4j/editor-model';
import type { EditorSurface } from '../../view/surface.js';
import { useEditorState, selectSource } from '../state.js';
import { useT } from '../i18n.js';

interface Row { address: string; style: string; text: string; kind: 'p' | 'tbl'; sourceId: string | undefined }

/** One block's entries as the engine gave them, keyed by the block's element object. */
interface Memo { element: object; paragraphs: OutlineParagraph[]; tables: OutlineTable[] }

/** The entries of `outline` grouped by the top-level block each belongs to, by its address's first step. */
function groupByBlock(outline: Outline, elements: readonly object[]): Memo[] {
  const index = (address: string): number => Number(address.split('/')[1]);
  const out = elements.map((element): Memo => ({ element, paragraphs: [], tables: [] }));
  for (const p of outline.paragraphs) out[index(p.address)]?.paragraphs.push(p);
  for (const t of outline.tables) out[index(t.address)]?.tables.push(t);
  return out;
}

/** The engine's outline of the body the session holds, whole or for some blocks, under one fresh context. */
async function readOutline(session: EditorDocument, blocks?: number[]): Promise<Outline> {
  const body = new Body(session.main, session.body, 'body', session.pkg);
  const context = await Texts.contextOf(session.pkg);
  return body.outline(undefined, context, blocks ? { blocks } : undefined);
}

export function OutlinePanel({ editor, session }: { editor: EditorSurface; session: EditorDocument }) {
  const tr = useT();
  const state = useEditorState(editor);
  const [memos, setMemos] = useState<Memo[]>([]);

  useEffect(() => {
    let current = true;
    const elements = (session.body.content ?? []) as object[];
    const run = async (): Promise<void> => {
      const same = memos.length === elements.length && memos.every((m, i) => m.element === elements[i]);
      if (same) return;
      const changed: number[] = [];
      if (memos.length === elements.length) elements.forEach((el, i) => { if (memos[i]!.element !== el) changed.push(i); });
      const partial = memos.length === elements.length && changed.length > 0 && changed.length < elements.length && !memos.some((m) => m.paragraphs.some((p) => p.joined && p.joined.length > 0));
      let next: Memo[];
      if (partial) {
        const read = groupByBlock(await readOutline(session, changed), elements);
        next = memos.map((m, i) => (changed.includes(i) ? read[i]! : m));
      } else {
        next = groupByBlock(await readOutline(session), elements);
      }
      if (current) setMemos(next);
    };
    void run().catch(() => undefined);
    return () => { current = false; };
    // The body's block list is compared on every transaction (identity per block, cheap); `memos` is read, not a trigger.
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [session, state]);

  const byElement = new Map<object, string>();
  session.doc.descendants((node) => {
    const id = node.attrs['sourceId'] as string | null;
    const el = id ? session.ctx.ids.get(id) : undefined;
    if (el && id) byElement.set(el.value as object, id);
    return true;
  });
  const body = new Body(session.main, session.body, 'body', session.pkg);
  const rows: Row[] = [];
  for (const m of memos) {
    for (const p of m.paragraphs) {
      const el = body.elementAt(p.address);
      rows.push({ address: p.address, style: p.style, text: p.text, kind: 'p', sourceId: el ? byElement.get(el.element.value as object) : undefined });
    }
    for (const t of m.tables) {
      const el = body.elementAt(t.address);
      rows.push({ address: t.address, style: tr('outline.table'), text: `${t.rows ?? ''}`, kind: 'tbl', sourceId: el ? byElement.get(el.element.value as object) : undefined });
    }
  }
  rows.sort((a, b) => a.address.localeCompare(b.address, undefined, { numeric: true }));

  return (
    <section className="panel outline-panel" aria-label={tr('tab.outline')} data-testid="panel-outline">
      <h2>{tr('tab.outline')}</h2>
      <div className="table-box">
      <table>
        <thead><tr><th>{tr('outline.address')}</th><th>{tr('outline.style')}</th><th>{tr('outline.text')}</th></tr></thead>
        <tbody>
          {rows.map((r) => (
            <tr key={r.address} className={r.kind}>
              <td><button type="button" className="link" onClick={() => { if (r.sourceId) selectSource(editor, r.sourceId); }} disabled={!r.sourceId}>{r.address}</button></td>
              <td>{r.style}</td>
              <td className="text">{r.text.slice(0, 80)}</td>
            </tr>
          ))}
        </tbody>
      </table>
      </div>
    </section>
  );
}

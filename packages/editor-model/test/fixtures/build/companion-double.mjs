// SPDX-License-Identifier: GPL-3.0-or-later
//
// A double for the companion's `paginate` (ED-003 section 6.5): the contract the editor speaks,
// answered by a Node function that moves every `w:lastRenderedPageBreak` marker one paragraph
// down (and drops one that has nowhere to go), which is enough to prove the reload, the labels
// and the undo without docx4j's layout. `paginateDouble(bytes)` is the tool; `companionRoutes`
// serves the contract on any Node request handler (Vite's preview server for the browser tests):
//
//   GET  <base>/           -> { "tools": ["paginate"] }
//   POST <base>/paginate   -> the docx bytes in, the docx bytes out
//
// The real companion is E4's `@docx4j/editor-mcp` over docx4j-mcp's `paginate` tool.
import { WordprocessingMLPackage, typeNameOf } from '@docx4j/core-ts';

/** The markers moved one paragraph down; the output re-marshalled by the engine. */
export async function paginateDouble(bytes) {
  const pkg = await WordprocessingMLPackage.load(bytes);
  const main = pkg.getMainDocumentPart();
  const document = await main.getContents();
  const paragraphs = (document.body?.content ?? []).filter((el) => typeNameOf(el) === 'org_docx4j_wml.P');
  const isMarker = (item) => item.name.localPart === 'lastRenderedPageBreak';
  const moves = [];
  paragraphs.forEach((p, i) => {
    for (const run of p.value.content ?? []) {
      if (typeNameOf(run) !== 'org_docx4j_wml.R') continue;
      const items = run.value.content ?? [];
      const at = items.findIndex(isMarker);
      if (at >= 0) { const [marker] = items.splice(at, 1); moves.push({ marker, to: i + 1 }); }
    }
  });
  for (const { marker, to } of moves) {
    const target = paragraphs[to];
    if (!target) continue;   // nowhere to go: dropped
    const firstRun = (target.value.content ?? []).find((el) => typeNameOf(el) === 'org_docx4j_wml.R');
    if (firstRun) (firstRun.value.content ??= []).unshift(marker);
    else (target.value.content ??= []).unshift({ name: { namespaceURI: 'http://schemas.openxmlformats.org/wordprocessingml/2006/main', localPart: 'r' }, value: { TYPE_NAME: 'org_docx4j_wml.R', content: [marker] } });
  }
  return pkg.save();
}

/** The contract on a Node `(req, res, next)` handler under `base` (`/companion`), for Vite's dev and preview servers. */
export function companionRoutes(base = '/companion') {
  return (req, res, next) => {
    const url = (req.url ?? '').split('?')[0];
    if (url === base || url === `${base}/`) {
      if (req.method !== 'GET') { res.statusCode = 405; res.end(); return; }
      res.setHeader('Content-Type', 'application/json');
      res.setHeader('Access-Control-Allow-Origin', '*');
      res.end(JSON.stringify({ tools: ['paginate'] }));
      return;
    }
    if (url === `${base}/paginate`) {
      if (req.method !== 'POST') { res.statusCode = 405; res.end(); return; }
      const chunks = [];
      req.on('data', (c) => chunks.push(c));
      req.on('end', () => {
        paginateDouble(new Uint8Array(Buffer.concat(chunks))).then((out) => {
          res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.wordprocessingml.document');
          res.setHeader('Access-Control-Allow-Origin', '*');
          res.end(Buffer.from(out));
        }, (e) => { res.statusCode = 500; res.end(String(e)); });
      });
      return;
    }
    next();
  };
}

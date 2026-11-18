// SPDX-License-Identifier: GPL-3.0-or-later
//
// The Java oracle (ED-003 section 4.8 as amended, step B2): core-ts's `Docx4J.bind` harness
// writes one JSON golden per template-and-data pair, naming every control by `w:id` with what
// happened to it; `test/fixtures/opendope/refresh.mjs` copies the goldens here. This test
// compares them with the verdicts of section 4.7 over the same template and data. Until the
// harness exists (core-ts CR-005), there are no goldens and every pair skips with a notice.
//
// The golden's shape, as the editor asked core-ts for it:
//   { "template": "invoice.docx", "data": "invoice-data.xml", "docx4j": "17.2.0",
//     "controls": { "<w:id>": { "value"?: string, "kept"?: boolean, "count"?: number, "unchanged"?: true } } }
import { describe, test, expect } from 'vitest';
import { readFile, stat } from 'node:fs/promises';
import { basename, join } from 'node:path';
import { EditorDocument, listControls, previewInstance } from '../src/index.mjs';
import { fixturesDir } from './helpers.mjs';

interface Golden {
  controls: Record<string, { value?: string; kept?: boolean; count?: number; unchanged?: true }>;
}

interface Pair {
  template: string;
  data?: string;
}

const dir = join(fixturesDir, 'opendope');
const manifest = JSON.parse(await readFile(join(dir, 'manifest.json'), 'utf8')) as { pairs: Pair[] };

async function goldenFor(pair: Pair): Promise<Golden | undefined> {
  const path = join(dir, 'goldens', `${basename(pair.template, '.docx')}.json`);
  try {
    await stat(path);
  } catch {
    return undefined;
  }
  return JSON.parse(await readFile(path, 'utf8')) as Golden;
}

describe.each(manifest.pairs)('$template', (pair) => {
  test('the verdicts agree with docx4j\'s golden', async (ctx) => {
    const golden = await goldenFor(pair);
    if (!golden) {
      console.log(`oracle: no golden for ${pair.template} (run test/fixtures/opendope/refresh.mjs once core-ts's bind harness has produced one); skipped`);
      ctx.skip();
      return;
    }
    const session = await EditorDocument.open(new Uint8Array(await readFile(join(dir, pair.template))));
    if (pair.data) {
      // Step 1 of the pipeline: the data part replaced by the runtime XML.
      await session.template.load();
      const data = session.template.dataPart.part;
      expect(data, 'the data part must be determined to insert the runtime XML (REQ-004)').toBeDefined();
      data!.setXml(await readFile(join(dir, pair.data), 'utf8'));
      session.template.invalidate();
    } else {
      await session.template.load();
    }
    const t = session.template;
    for (const control of listControls(session.doc, session.controlEntry)) {
      const id = control.reading.id;
      if (id === undefined) continue;
      const expected = golden.controls[String(id)];
      if (!expected || expected.unchanged) continue;
      const verdict = t.verdictFor(control.sdtPr);
      if (expected.kept !== undefined) expect(verdict.kept, `condition on control ${id}`).toBe(expected.kept);
      if (expected.count !== undefined) expect(verdict.count, `repeat on control ${id}`).toBe(expected.count);
      if (expected.value !== undefined) expect(verdict.value ?? '', `value of control ${id}`).toBe(expected.value);
    }
    // The preview (section 4.7, step B5): the text the bind step put into each valued control agrees with the golden's value.
    const preview = await previewInstance(session);
    for (const control of listControls(preview.doc, preview.controls)) {
      const id = control.reading.id;
      if (id === undefined || control.pos === undefined) continue;
      const expected = golden.controls[String(id)];
      if (!expected || expected.value === undefined) continue;
      expect(preview.doc.nodeAt(control.pos)!.textContent, `bound text of control ${id}`).toBe(expected.value);
    }
  });
});

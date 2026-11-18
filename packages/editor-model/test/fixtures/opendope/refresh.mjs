// SPDX-License-Identifier: GPL-3.0-or-later
//
// Refresh the OpenDoPE fixtures (ED-003 section 4.8 as amended, step B2): the templates and data
// files come from docx4j's databinding samples, the goldens from core-ts's `Docx4J.bind` harness
// when it exists. Copied, not referenced (ED-002 section 3.8), so the package's tests run from
// npm. Run from anywhere: `node packages/editor-model/test/fixtures/opendope/refresh.mjs`.
import { copyFile, mkdir, readdir, stat } from 'node:fs/promises';
import { dirname, join } from 'node:path';
import { fileURLToPath } from 'node:url';

const here = dirname(fileURLToPath(import.meta.url));
const repos = join(here, '..', '..', '..', '..', '..', '..');
const samples = join(repos, 'docx4j', 'docx4j-samples-docx4j', 'sample-docs', 'databinding');
/** Where core-ts's bind harness writes its goldens (its `test/golden/` holds what docx4j answers, CR-005 phase B; the editor copies them). */
const goldens = join(repos, 'docx4j-core-ts', 'test', 'golden', 'opendope');

/** The templates copied here; `invoice.docx` and `invoice2013.docx` are root fixtures already and are not duplicated. */
const templates = ['binding-simple.docx', 'conventions.docx', 'non-element-bind.docx', 'hyperlink-binding-test.docx', 'picture.docx', 'invoice_Saxon_XPath2.docx'];
const data = ['binding-simple-data.xml', 'invoice-data.xml', 'invoice2013.xml'];

async function exists(path) {
  try { await stat(path); return true; } catch { return false; }
}

for (const name of [...templates, ...data]) {
  const from = join(samples, name);
  if (!(await exists(from))) { console.warn(`missing in docx4j: ${from}`); continue; }
  await copyFile(from, join(here, name));
  console.log(`copied ${name}`);
}
if (await exists(goldens)) {
  await mkdir(join(here, 'goldens'), { recursive: true });
  for (const name of await readdir(goldens)) {
    if (!name.endsWith('.json')) continue;
    await copyFile(join(goldens, name), join(here, 'goldens', name));
    console.log(`copied golden ${name}`);
  }
} else {
  console.log(`no goldens yet at ${goldens} (core-ts CR-005's bind harness); the oracle test skips with a notice`);
}

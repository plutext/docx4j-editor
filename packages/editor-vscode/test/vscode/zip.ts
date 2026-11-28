// SPDX-License-Identifier: GPL-3.0-or-later
//
// A package's parts, for the round trip: every part the same as the source's but `[Content_Types].xml` and the `.rels`
// parts, which the engine regenerates by design (ED-002 section 3.7), as the round-trip suites compare.
import { readFileSync, writeFileSync } from 'node:fs';
import { strToU8, unzipSync, zipSync, strFromU8 } from 'fflate';

export function parts(path: string): Map<string, Uint8Array> {
  return new Map(Object.entries(unzipSync(new Uint8Array(readFileSync(path)))));
}

const regenerated = (name: string): boolean => name === '[Content_Types].xml' || name.endsWith('.rels');

/** The parts that differ, or are only on one side, `[Content_Types].xml` and the `.rels` parts aside. */
export function partsThatDiffer(source: string, saved: string): string[] {
  const a = parts(source);
  const b = parts(saved);
  const names = new Set([...a.keys(), ...b.keys()]);
  const out: string[] = [];
  for (const name of names) {
    if (regenerated(name)) continue;
    const x = a.get(name);
    const y = b.get(name);
    if (!x || !y || x.byteLength !== y.byteLength || x.some((v, i) => v !== y[i])) out.push(name);
  }
  return out.sort();
}

/** The main part's name: `word/document.xml`, or Word Online's `word/document22.xml`. */
function mainName(all: Map<string, Uint8Array> | Record<string, Uint8Array>): string {
  const names = all instanceof Map ? [...all.keys()] : Object.keys(all);
  return names.find((n) => /^word\/document\d*\.xml$/.test(n))!;
}

export function mainPart(path: string): string {
  const all = parts(path);
  return strFromU8(all.get(mainName(all))!);
}

/** A copy of a package with a paragraph holding a link to `target` before the body's last section, and its relationship. */
export function withLink(source: string, to: string, target: string, text: string): void {
  const all = Object.fromEntries(parts(source));
  const main = mainName(all);
  const relsName = main.replace(/^word\//, 'word/_rels/') + '.rels';
  const doc = strFromU8(all[main]!);
  const at = doc.lastIndexOf('<w:sectPr');
  const p = `<w:p><w:hyperlink r:id="rIdSuiteLink"><w:r><w:t>${text}</w:t></w:r></w:hyperlink></w:p>`;
  all[main] = strToU8(doc.slice(0, at) + p + doc.slice(at));
  const rels = strFromU8(all[relsName]!);
  const escaped = target.replace(/&/g, '&amp;').replace(/"/g, '&quot;').replace(/</g, '&lt;');
  all[relsName] = strToU8(rels.replace('</Relationships>', `<Relationship Id="rIdSuiteLink" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" Target="${escaped}" TargetMode="External"/></Relationships>`));
  writeFileSync(to, zipSync(all));
}

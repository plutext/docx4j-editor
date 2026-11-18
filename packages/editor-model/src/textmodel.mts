// SPDX-License-Identifier: GPL-3.0-or-later
//
// The package as the engine's text model reads it (ED-008 section 3.1). The text model (core-ts's `Texts`,
// CR-012) reads every part with `readContents()`: the live tree where the part is unmarshalled, else a private
// tree from the part's bytes. A session's main part holds the session's tree only after a tree run (it is
// handed over then, and at a save that changed a block); before that, the part's bytes are the file as it was
// opened, and a reading would give the file's text, not the document's, which in a room is the person's
// typing too. So a reader is given the package with its main part reading the session's tree, and nothing
// else changed: the real part is never unmarshalled or handed the tree, so a save afterwards is byte for
// byte where it was before (ED-002 section 3.7).
import type { TextPackage, WordprocessingMLPackage } from '@docx4j/core-ts';
import type * as wml from '@docx4j/generated-objects-ts/modules/org_docx4j_wml';

/**
 * The members of a main part that give its contents, each answered here with the tree given, so that no
 * reader of the view reaches the part's own bytes (`getContents` would unmarshal them and make the part live).
 */
const CONTENTS = new Set<PropertyKey>(['readContents', 'getContents', 'contents']);

/**
 * `pkg` as a text-model package whose main part reads `document`. The package itself where its main part
 * already holds `document` (after a tree run). Everything but the main part's contents is the package's own:
 * the styles, the numbering, the story parts and the relationships.
 */
export function textPackageOver(pkg: WordprocessingMLPackage, document: wml.Document): TextPackage {
  const main = pkg.getMainDocumentPart();
  if (main.isUnmarshalled && main.contents === document) return pkg;
  const mainView = new Proxy(main, {
    get(target, prop) {
      if (CONTENTS.has(prop)) return prop === 'contents' ? document : () => Promise.resolve(document);
      const value: unknown = Reflect.get(target, prop, target);
      return typeof value === 'function' ? (value as (...args: unknown[]) => unknown).bind(target) : value;
    },
  });
  return new Proxy(pkg, {
    get(target, prop) {
      if (prop === 'getMainDocumentPart') return () => mainView;
      const value: unknown = Reflect.get(target, prop, target);
      return typeof value === 'function' ? (value as (...args: unknown[]) => unknown).bind(target) : value;
    },
  });
}

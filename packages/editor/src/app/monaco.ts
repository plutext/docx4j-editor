// SPDX-License-Identifier: GPL-3.0-or-later
//
// The application's one way to Monaco (ED-014 section 4.2, P14): the editor's features and the three languages the
// application shows, not the whole package, whose `monaco-editor` entry registers 81 languages and four language
// services. The console, the XML editors and the console's declarations import Monaco from here alone, so the workers
// are always set before an editor starts (`monaco-env.ts`). `test/monaco-imports.test.ts` holds the rule: a single
// `import … from 'monaco-editor'` elsewhere would bring the other languages back unseen.
//
// The editor's features, as the whole package has them: find, folding, Format, the menus, suggest, hover.
import 'monaco-editor/esm/vs/editor/edcore.main.js';
// The languages the application shows: the XML editors', and the console's TypeScript, with JavaScript beside it.
// The basic languages colour the text; the TypeScript contribution is the language service, and fills
// `monaco.languages.typescript`.
import 'monaco-editor/esm/vs/basic-languages/xml/xml.contribution.js';
import 'monaco-editor/esm/vs/basic-languages/typescript/typescript.contribution.js';
import 'monaco-editor/esm/vs/basic-languages/javascript/javascript.contribution.js';
import 'monaco-editor/esm/vs/language/typescript/monaco.contribution.js';
import './monaco-env.js';
export * from 'monaco-editor/esm/vs/editor/editor.api.js';

// SPDX-License-Identifier: GPL-3.0-or-later
//
// Monaco's workers (ED-003 section 7.3; ED-005 section 3.8): the TypeScript worker for the
// console, the editor worker for everything else (the XML editors have no language worker). Set
// once, by whichever of the console and the XML editors loads Monaco first; both are chunks of
// their own, so the Edit perspective's bundle never carries Monaco.
import editorWorker from 'monaco-editor/esm/vs/editor/editor.worker?worker';
import tsWorker from 'monaco-editor/esm/vs/language/typescript/ts.worker?worker';

(self as unknown as { MonacoEnvironment?: unknown }).MonacoEnvironment = {
  getWorker(_: unknown, label: string) {
    if (label === 'typescript' || label === 'javascript') return new tsWorker();
    return new editorWorker();
  },
};

// SPDX-License-Identifier: GPL-3.0-or-later
//
// The console (ED-001 section 3.5; ED-002 section 6.3; ED-003 section 7.3): Monaco in TypeScript
// mode with the engine's declarations, Run (Ctrl+Enter) through `runScript`, output and errors
// with their line, the last script kept. Pasted Office JS runs through the `Word` binding; before
// a run, the members the subset lacks are printed from Monaco's own diagnostics and the
// portability lint says which reads Word would refuse, with the load calls offered. Loaded with
// import() when the Developer perspective first shows it, so the Edit perspective's bundle stays
// what it is. Devtools-level trust: nothing runs without Run.
import { useCallback, useEffect, useRef, useState } from 'react';
import * as monaco from '../monaco.js';
import type { EditorDocument } from '@docx4j/editor-model';
import type { DocxEditorView } from '../../view/DocxEditorView.js';
import { runScript, editorBinding, type RunOutcome } from '../console-run.js';
import { loadDeclarations, loadAllModuleDeclarations, MODULE_COUNT } from '../declarations.js';
import { lintOfficeJs, type LintResult } from '../office-js-lint.js';
import type { ConsoleScript } from '../perspectives.js';
import { AskBox, type AskBoxProps } from './AskBox.js';
import { Fold } from './Fold.js';
import { t, useT } from '../i18n.js';

const LAST_SCRIPT = 'docx4j-editor.console.script';
const DEFAULT_SCRIPT = `// The document is on screen; the tree is here. Ctrl+Enter runs.
// pkg, body, main: the package and the content API (Office JS shapes).
// selection: the selected paragraphs, range and font; el and wml: the builders.
// Word: an add-in's Word.run(async (context) => { ... }) runs unchanged.
const heading = body.insertParagraph('Hello from the console', 'End');
heading.style = 'Heading 1';
return body.paragraphs.length;
`;

interface Line {
  kind: 'log' | 'warn' | 'error' | 'result';
  text: string;
}

function show(v: unknown): string {
  if (typeof v === 'string') return v;
  try { return JSON.stringify(v, (k, x) => (k === 'PARENT' ? undefined : x), 2) ?? String(v); } catch { return String(v); }
}

/** The codes the console's compiler options already ignore (top-level await). */
const IGNORED_CODES = new Set([1375, 1378]);

/**
 * The TypeScript worker's semantic diagnostics as lines: a missing member (2339: no such
 * property; 2551: with a suggestion) is "not in the subset"; anything else is a type error, said
 * as Monaco says it (Jason, 2026-09-25: a `context` used outside its callback drew a squiggle
 * that Check called nothing to report).
 */
async function typeDiagnostics(model: monaco.editor.ITextModel): Promise<string[]> {
  const worker = await monaco.languages.typescript.getTypeScriptWorker();
  const client = await worker(model.uri);
  const diagnostics = await client.getSemanticDiagnostics(model.uri.toString());
  // The type's text is the shimmed view's, pages long: the member and the class's first words are what the line needs.
  const flatten = (m: string | monaco.languages.typescript.DiagnosticMessageChain): string => (typeof m === 'string' ? m : m.messageText).replace(/on type '([^']{0,60})[^']*'/, (_, head: string, offset: number, whole: string) => (whole.length > offset + head.length + 10 ? `on type '${head}…'` : `on type '${head}'`));
  return diagnostics
    .filter((d) => !IGNORED_CODES.has(d.code))
    .map((d) => {
      const line = d.start !== undefined ? model.getPositionAt(d.start).lineNumber : '?';
      const subset = d.code === 2339 || d.code === 2551;
      return t(subset ? 'console.diagnostic.subset' : 'console.diagnostic.type', { line, message: flatten(d.messageText) });
    });
}

export interface ConsoleProps {
  editor: DocxEditorView;
  session: EditorDocument;
  perspective: { id: string; panel: { open: boolean; tab: string } };
  /** A script handed over: the reveal-codes pane's replaces the editor's text, the palette's command line is inserted at the caret. */
  script?: ConsoleScript | null;
  /** Save the editor's script as a macro (ED-003 section 7.4): the shell asks the name and the home. */
  onSaveMacro?: (script: string) => void;
  /** The ask mode's shell callbacks (ED-003 section 7.5); absent hides the box. */
  ask?: Omit<AskBoxProps, 'session' | 'editor' | 'onScript'>;
}

export default function Console({ editor, session, perspective, script, onSaveMacro, ask }: ConsoleProps) {
  useT();
  const host = useRef<HTMLDivElement>(null);
  const model = useRef<monaco.editor.IStandaloneCodeEditor | null>(null);
  const [lines, setLines] = useState<Line[]>([]);
  const [running, setRunning] = useState(false);
  const [types, setTypes] = useState<'loading' | 'ready' | 'failed'>('loading');
  const [allTypes, setAllTypes] = useState<'no' | 'loading' | 'yes'>('no');
  const [lint, setLint] = useState<LintResult | null>(null);
  /** Bumped by Clear output: the ask box drops its request and answer too (Jason, 2026-09-25: Clear left them). */
  const [cleared, setCleared] = useState(0);

  /** The hand-over applied when the editor was created, so the effect below does not apply it again. */
  const applied = useRef<number | null>(null);
  useEffect(() => {
    if (!host.current) return;
    let initial = DEFAULT_SCRIPT;
    try { initial = localStorage.getItem(LAST_SCRIPT) ?? DEFAULT_SCRIPT; } catch { /* ignore */ }
    if (script) {
      initial = script.mode === 'replace' ? script.text : `${initial.trimEnd()}\n${script.text}\n`;
      applied.current = script.seq;
    }
    // A model with a `.ts` name: the worker takes a file with no extension for JavaScript when
    // `allowJs` is on (ED-003 section 11.15 item 5) and then reports no type errors at all.
    const textModel = monaco.editor.createModel(initial, 'typescript', monaco.Uri.parse('file:///console/script.ts'));
    const instance = monaco.editor.create(host.current, {
      model: textModel,
      minimap: { enabled: false },
      automaticLayout: true,
      fontSize: 13,
      lineNumbers: 'on',
      scrollBeyondLastLine: false,
      ariaLabel: t('console.script'),
    });
    model.current = instance;
    // The live editor for the browser tests, as `window.docx4jEditor` is the view: a script set through it is not retyped through Monaco's auto-closing.
    (window as unknown as { docx4jConsole?: monaco.editor.IStandaloneCodeEditor }).docx4jConsole = instance;
    (window as unknown as { docx4jMonaco?: typeof monaco }).docx4jMonaco = monaco;
    void loadDeclarations().then(() => setTypes('ready')).catch(() => setTypes('failed'));
    return () => { instance.dispose(); textModel.dispose(); model.current = null; delete (window as unknown as { docx4jConsole?: unknown }).docx4jConsole; };
    // The editor is created once; a handed-over script is applied by the effect below.
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, []);

  const setScript = useCallback((text: string) => { model.current?.setValue(text); }, []);

  useEffect(() => {
    const instance = model.current;
    if (!script || !instance || applied.current === script.seq) return;
    applied.current = script.seq;
    if (script.mode === 'replace') { instance.setValue(script.text); return; }
    // Inserted at the caret (a whole line when the caret is not at a line's start), the caret after it.
    const sel = instance.getSelection() ?? new monaco.Selection(1, 1, 1, 1);
    const atLineStart = sel.startColumn === 1;
    instance.executeEdits('palette', [{ range: sel, text: `${atLineStart ? '' : '\n'}${script.text}\n`, forceMoveMarkers: true }]);
    instance.focus();
  }, [script]);

  /** The pre-run pass (ED-003 section 7.3): the subset's missing members from Monaco, the portability lint; both printed, nothing run. */
  const check = useCallback(async (push: (kind: Line['kind'], text: string) => void): Promise<void> => {
    const instance = model.current;
    if (!instance) return;
    const code = instance.getValue();
    const textModel = instance.getModel();
    if (textModel && types === 'ready') {
      try { for (const d of await typeDiagnostics(textModel)) push('warn', d); } catch { /* the worker is not there: the run's NotSupportedError is the net */ }
    }
    const result = lintOfficeJs(code);
    setLint(result.warnings.length > 0 ? result : null);
    for (const w of result.warnings) push('warn', w.message);
  }, [types]);

  const insertLoads = useCallback(() => {
    const instance = model.current;
    if (!instance || !lint?.insertion) return;
    instance.executeEdits('lint', [{ range: new monaco.Range(lint.insertion.line, 1, lint.insertion.line, 1), text: lint.insertion.text, forceMoveMarkers: true }]);
    const again = lintOfficeJs(instance.getValue());
    setLint(again.warnings.length > 0 ? again : null);
    instance.focus();
  }, [lint]);

  const run = useCallback(async () => {
    const instance = model.current;
    if (!instance || running) return;
    const code = instance.getValue();
    try { localStorage.setItem(LAST_SCRIPT, code); } catch { /* ignore */ }
    setRunning(true);
    const out: Line[] = [];
    const push = (kind: Line['kind'], args: unknown[]): void => { out.push({ kind, text: args.map(show).join(' ') }); setLines([...out]); };
    await check((kind, text) => push(kind, [text]));
    let outcome: RunOutcome;
    try {
      outcome = await runScript(session, editorBinding(editor), perspective, code, {
        log: (...a) => push('log', a), warn: (...a) => push('warn', a), error: (...a) => push('error', a),
      });
    } catch (e) {
      outcome = { error: e, changed: false };
    }
    if (outcome.error !== undefined) {
      const error = `${(outcome.error as Error).name ?? 'Error'}: ${(outcome.error as Error).message ?? String(outcome.error)}`;
      push('error', [outcome.line !== undefined ? t('console.lineOf', { error, line: outcome.line }) : error]);
    } else if (outcome.value !== undefined) push('result', [outcome.value]);
    else push('result', [t('console.noValue')]);
    setRunning(false);
    editor.focus();
  }, [editor, session, perspective, running, check]);

  const checkOnly = useCallback(async () => {
    const out: Line[] = [];
    const push = (kind: Line['kind'], text: string): void => { out.push({ kind, text }); setLines([...out]); };
    await check(push);
    if (out.length === 0) push('result', t('console.nothingToReport'));
  }, [check]);

  useEffect(() => {
    const instance = model.current;
    if (!instance) return;
    const disposable = instance.addAction({ id: 'docx4j.run', label: t('console.run.action'), keybindings: [monaco.KeyMod.CtrlCmd | monaco.KeyCode.Enter], run: () => { void run(); } });
    return () => disposable.dispose();
  }, [run]);

  const loadAll = useCallback(() => {
    setAllTypes('loading');
    void loadAllModuleDeclarations().then(() => setAllTypes('yes')).catch(() => setAllTypes('no'));
  }, []);

  return (
    <Fold id="console" className="panel console-panel" label={t('tab.console')} testId="panel-console" title={t('tab.console')}
      meta={<span data-testid="console-types">{types === 'loading' ? t('console.types.loading') : types === 'ready' ? (allTypes === 'yes' ? t('console.types.all', { count: MODULE_COUNT + 1 }) : t('console.types.ready')) : t('console.types.unavailable')}</span>}>
      <div ref={host} className="console-editor" data-testid="console-editor" />
      <div className="pane-actions">
        <button type="button" data-testid="console-run" onClick={() => void run()} disabled={running}>{t('console.run')}</button>
        <button type="button" data-testid="console-check" title={t('console.check.title')} onClick={() => void checkOnly()}>{t('console.check')}</button>
        {lint?.insertion && <button type="button" data-testid="console-insert-loads" title={t('console.insertLoads.title')} onClick={insertLoads}>{t('console.insertLoads')}</button>}
        {onSaveMacro && <button type="button" data-testid="console-save-macro" onClick={() => { const v = model.current?.getValue(); if (v !== undefined) onSaveMacro(v); }}>{t('console.saveMacro')}</button>}
        <button type="button" data-testid="console-clear" title={t('console.clear.title')} onClick={() => { setLines([]); setLint(null); setCleared((n) => n + 1); }}>{t('console.clear')}</button>
        <button type="button" data-testid="console-clear-script" title={t('console.clearScript.title')} onClick={() => { model.current?.setValue(''); model.current?.focus(); }}>{t('console.clearScript')}</button>
        {/* Out of the heading: a button inside the fold's summary would take the click meant to fold it. */}
        {types === 'ready' && allTypes !== 'yes' && <button type="button" className="link" data-testid="console-load-all-types" onClick={loadAll} disabled={allTypes === 'loading'}>{allTypes === 'loading' ? t('console.loadAll.loading') : t('console.loadAll', { count: MODULE_COUNT })}</button>}
      </div>
      <ol className="console-output" aria-label={t('console.output')} data-testid="console-output">
        {lines.map((l, i) => <li key={i} className={l.kind}><pre>{l.text}</pre></li>)}
      </ol>
      {ask && <AskBox {...ask} session={session} editor={editor} onScript={setScript} reset={cleared} />}
    </Fold>
  );
}

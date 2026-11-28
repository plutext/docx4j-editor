// SPDX-License-Identifier: GPL-3.0-or-later
//
// XML in Monaco (ED-005 sections 3.8 and 4.3): the editor reveal codes, the Styles tab and the
// parts panel edit XML in, with Monaco's own find and replace (Ctrl+F and Ctrl+H inside it,
// regular expressions included), syntax colouring and folding. A problem is a marker at its line
// and column. Loaded with `import()` when first shown, as the console is. The live instance is
// on `window.docx4jXmlEditors` by test id, for the browser tests: text set through it is not
// retyped through Monaco's auto-closing, as `setConsoleScript` does for the console.
import { useEffect, useRef, useState } from 'react';
import * as monaco from '../monaco.js';
import { formatXml, foldingRangesOf, type XmlFocus } from '../xml.js';
import { t, useT } from '../i18n.js';

// Format Document (Shift+Alt+F, the context menu, the parts panel's Format) for every XML editor: `formatXml`, as one
// edit, so Ctrl+Z undoes it. A text that does not parse is left as it is; the parts panel says why.
monaco.languages.registerDocumentFormattingEditProvider('xml', {
  provideDocumentFormattingEdits(model) {
    try {
      const formatted = formatXml(model.getValue());
      return formatted === model.getValue() ? [] : [{ range: model.getFullModelRange(), text: formatted }];
    } catch {
      return [];
    }
  },
});

// Folding by element for every XML editor: a region from an element's start tag to its end tag, so a folded row or
// cell hides its end tag too; Monaco's own fallback folds by indentation, which leaves the end tag showing.
monaco.languages.registerFoldingRangeProvider('xml', {
  provideFoldingRanges(model) {
    return foldingRangesOf(model.getValue());
  },
});

/** An entry a panel adds to the editor's context menu (the Data tab's Bind and Copy XPath): given the caret's offset and the text. */
export interface XmlContextAction {
  id: string;
  label: string;
  run(offset: number, text: string): void;
}

/** What a panel can ask of its XML editor. */
export interface XmlEditorApi {
  /** Format Document: the text indented as one edit, which Ctrl+Z undoes. */
  format(): Promise<void>;
}

export interface XmlProblem {
  line: number;
  column?: number;
  message: string;
}

export interface XmlCodeEditorProps {
  value: string;
  onChange?: (value: string) => void;
  readOnly?: boolean;
  /** The problems to mark: a parse error at its line, a check's findings at theirs. */
  problems?: readonly XmlProblem[];
  ariaLabel: string;
  testId: string;
  /** Called with the editor's API once it is made, and with null when it goes. */
  api?: (api: XmlEditorApi | null) => void;
  /** A smaller box, for a short text (a part's relationships). */
  compact?: boolean;
  /** Entries for the context menu, first in it; fixed when the editor is made. */
  actions?: readonly XmlContextAction[];
  /**
   * The lines to mark and the regions to fold around them (reveal codes in a table: the caret's cell, the other rows
   * and cells folded). Null clears the mark; undefined leaves the editor as it is (while a text is being edited).
   */
  focus?: XmlFocus | null;
}

let sequence = 0;

/** Where the caret is, and the character after it: its code point, and the character itself when it draws something. */
export interface CaretReading {
  line: number;
  column: number;
  codePoint?: number;
  char?: string;
}

/** The caret's reading in a model (the code point of the character after the caret; a blank or an invisible character is a code point alone). */
export function caretReading(model: monaco.editor.ITextModel, position: monaco.IPosition): CaretReading {
  const text = model.getLineContent(position.lineNumber);
  const index = position.column - 1;
  if (index >= text.length) return { line: position.lineNumber, column: position.column };
  const codePoint = text.codePointAt(index)!;
  const char = String.fromCodePoint(codePoint);
  return { line: position.lineNumber, column: position.column, codePoint, ...(/[\p{L}\p{N}\p{P}\p{S}]/u.test(char) ? { char } : {}) };
}

/** `U+200B` for a code point. */
export const codePointLabel = (codePoint: number): string => `U+${codePoint.toString(16).toUpperCase().padStart(4, '0')}`;

export default function XmlCodeEditor({ value, onChange, readOnly = false, problems = [], ariaLabel, testId, api, compact = false, actions = [], focus }: XmlCodeEditorProps) {
  const host = useRef<HTMLDivElement>(null);
  const editor = useRef<monaco.editor.IStandaloneCodeEditor | null>(null);
  const marks = useRef<monaco.editor.IEditorDecorationsCollection | null>(null);
  /** The caret's line, column and the code point after it (Jason, 2026-10-15: handy for a U+200B), shown under the editor. */
  const [caret, setCaret] = useState<CaretReading | null>(null);
  useT();
  const changed = useRef(onChange);
  changed.current = onChange;

  useEffect(() => {
    if (!host.current) return;
    const model = monaco.editor.createModel(value, 'xml', monaco.Uri.parse(`file:///xml/${testId}-${++sequence}.xml`));
    const instance = monaco.editor.create(host.current, {
      model,
      readOnly,
      minimap: { enabled: false },
      automaticLayout: true,
      fontSize: 12,
      lineNumbers: 'on',
      folding: true,
      foldingMaximumRegions: 50000,   // a large table's rows and cells, each element a region
      wordWrap: 'on',
      scrollBeyondLastLine: false,
      ariaLabel,
    });
    editor.current = instance;
    const sub = model.onDidChangeContent(() => changed.current?.(model.getValue()));
    const read = (): void => { const at = instance.getPosition(); setCaret(at ? caretReading(model, at) : null); };
    const cursorSub = instance.onDidChangeCursorPosition(read);
    const contentSub = model.onDidChangeContent(read);
    read();
    const registry = ((window as unknown as { docx4jXmlEditors?: Record<string, monaco.editor.IStandaloneCodeEditor> }).docx4jXmlEditors ??= {});
    registry[testId] = instance;
    api?.({ format: async () => { await instance.getAction('editor.action.formatDocument')?.run(); } });
    const added = actions.map((a, i) => instance.addAction({
      id: `docx4j.${a.id}`,
      label: a.label,
      contextMenuGroupId: '0_docx4j',
      contextMenuOrder: i,
      run: (ed) => {
        const m = ed.getModel();
        const at = ed.getPosition();
        if (m && at) a.run(m.getOffsetAt(at), m.getValue());
      },
    }));
    return () => {
      for (const d of added) d.dispose();
      api?.(null);
      sub.dispose();
      cursorSub.dispose();
      contentSub.dispose();
      if (registry[testId] === instance) delete registry[testId];
      instance.dispose();
      model.dispose();
      editor.current = null;
      marks.current = null;
    };
    // The editor is made once; the value and the options follow below.
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, []);

  // A value set from outside (another block under the caret, a Revert) replaces the text.
  useEffect(() => {
    const instance = editor.current;
    if (instance && instance.getValue() !== value) instance.setValue(value);
  }, [value]);

  useEffect(() => { editor.current?.updateOptions({ readOnly }); }, [readOnly]);

  // The focus: the marked lines, then everything around them folded and the lines brought into view. The fold actions
  // wait for Monaco's folding model, which follows the text set above; neither takes the focus from the document.
  useEffect(() => {
    const instance = editor.current;
    if (!instance || focus === undefined) return;
    (marks.current ??= instance.createDecorationsCollection()).set(focus ? [{
      range: new monaco.Range(focus.line, 1, focus.endLine, 1),
      options: { isWholeLine: true, className: 'xml-focus', linesDecorationsClassName: 'xml-focus-gutter' },
    }] : []);
    if (!focus) return;
    let cancelled = false;
    void (async () => {
      instance.setPosition({ lineNumber: focus.line, column: 1 });
      await instance.getAction('editor.unfoldAll')?.run();
      if (cancelled) return;
      if (focus.fold.length > 0) await instance.getAction('editor.fold')?.run({ selectionLines: focus.fold.map((l) => l - 1) });
      if (cancelled) return;
      instance.revealLineNearTop(focus.line);
    })();
    return () => { cancelled = true; };
  }, [focus]);

  useEffect(() => {
    const model = editor.current?.getModel();
    if (!model) return;
    monaco.editor.setModelMarkers(model, 'docx4j', problems.map((p) => ({
      severity: monaco.MarkerSeverity.Error,
      message: p.message,
      startLineNumber: p.line,
      startColumn: p.column ?? 1,
      endLineNumber: p.line,
      endColumn: p.column !== undefined ? p.column + 1 : model.getLineMaxColumn(Math.min(Math.max(1, p.line), model.getLineCount())),
    })));
  }, [problems]);

  const first = problems[0];
  return (
    <>
      <div ref={host} className={compact ? 'xml-code compact' : 'xml-code'} data-testid={testId} data-readonly={readOnly || undefined} aria-invalid={problems.length > 0} data-error-line={first ? first.line : undefined} data-focus={focus ? `${focus.line}-${focus.endLine}` : undefined} />
      {caret && (
        <div className="xml-caret" data-testid={`${testId}-caret`}>
          <span>{t('xml.caret', { line: String(caret.line), column: String(caret.column) })}</span>
          {caret.codePoint !== undefined && <span className="xml-caret-char">{codePointLabel(caret.codePoint)}{caret.char !== undefined ? ` \u201C${caret.char}\u201D` : ''}</span>}
        </div>
      )}
    </>
  );
}

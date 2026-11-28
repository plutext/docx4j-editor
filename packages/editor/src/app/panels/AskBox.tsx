// SPDX-License-Identifier: GPL-3.0-or-later
//
// The ask box (ED-001 section 3.5; ED-003 section 7.5): natural language in, one of three fenced
// forms out. The request is built from the console's context and shown before it is sent (ED-001
// section 11 item 5); the answer is content previewed through the projection with Insert,
// Replace and Copy, a script handed to the console editor and never run, or an explanation.
// Nothing touches the document until a button is pressed, which is the difference from E4's agent.
//
// In VS Code (ED-014 section 4.6, P44 to P46) the page sends nothing itself: the host sends to one of VS Code's models or,
// with the key VS Code keeps, to the Anthropic API. The model the person picks chooses the route; the chat route stays.
import { useCallback, useEffect, useRef, useState, useSyncExternalStore } from 'react';
import { toApiScript, selectedBranchOf, type Element } from '@docx4j/core-ts';
import type * as wml from '@docx4j/generated-objects-ts/modules/org_docx4j_wml';
import { projectBody, selectionSpec, selectionBinding, type EditorDocument } from '@docx4j/editor-model';
import type { DocxEditorView } from '../../view/DocxEditorView.js';
import { StoryView } from '../../view/StoryView.js';
import { buildContext, buildRequest, chatPrompt, type AskRequest } from '../ask-context.js';
import { anthropicProvider, askOption, setAskOption, describeError, wasCancelled, askChoice, setKeyCommand, ASK_MODELS, OPTION_ASK_KEY, OPTION_ASK_BASE_URL, OPTION_ASK_MODEL, OPTION_ASK_SEND_AT_ONCE, VSCODE_MODEL, type AskAnswer, type AskRoute } from '../ask.js';
import { host, AskFailure, type HostAsk, type HostModel } from '../host.js';
import { modelKey } from '../host-protocol.js';
import { parseAnswer, type AskForm } from '../ask-answer.js';
import { parseElements, blocksIn } from '../xml.js';
import { t, tNodes, useT, formatNumber } from '../i18n.js';
import { Fold } from './Fold.js';

export interface AskBoxProps {
  session: EditorDocument;
  editor: DocxEditorView;
  /** A script answer goes to the console editor, shown and not run. */
  onScript: (code: string) => void;
  onStatus: (text: string) => void;
  /** Bumped by the console's Clear output: the request, the answer and the error go. */
  reset?: number;
}

/** A request built and shown, with the route and the model's name it was built for. */
interface Pending {
  request: AskRequest;
  route: AskRoute;
  label: string;
}

const noKeyStored = (): boolean => false;
const noSubscription = (): (() => void) => () => undefined;

/** Whether the host keeps an Anthropic key (VS Code, P45): false in a browser, whose key is the page's own. */
export function useHostKeyStored(hostAsk: HostAsk | undefined = host().ask()): boolean {
  const subscribe = useCallback((listener: () => void) => (hostAsk ? hostAsk.onKeyStored(listener) : noSubscription()), [hostAsk]);
  const snapshot = useCallback(() => (hostAsk ? hostAsk.keyStored() : noKeyStored()), [hostAsk]);
  return useSyncExternalStore(subscribe, snapshot);
}

interface Answer {
  raw: AskAnswer;
  form: AskForm;
  /** The parsed blocks of a wml answer, or the parse error. */
  elements?: Element[];
  parseError?: string;
}

/** The answer's blocks rendered through the projection, read-only, with the document's stylesheet (the body view's, in the page). */
function AskPreview({ session, elements }: { session: EditorDocument; elements: Element[] }) {
  const ref = useRef<HTMLDivElement>(null);
  useEffect(() => {
    if (!ref.current) return;
    const { doc, ctx } = projectBody({ content: elements } as unknown as wml.Body, session.schema);
    const view = new StoryView(ref.current, session, { key: 'ask-preview', doc, ctx }, { label: t('ask.preview') });
    return () => view.destroy();
  }, [session, elements]);
  return <div ref={ref} className="ask-preview" data-testid="ask-preview" />;
}

export function AskBox({ session, editor, onScript, onStatus, reset = 0 }: AskBoxProps) {
  useT();
  const hostAsk = host().ask();
  const [prompt, setPrompt] = useState('');
  const [kept, setKept] = useState(() => askOption(OPTION_ASK_MODEL));
  const [sendAtOnce, setSendAtOnce] = useState(() => askOption(OPTION_ASK_SEND_AT_ONCE) === 'true');
  const [request, setRequest] = useState<Pending | null>(null);
  /** The name of the model a request is being sent to, while it is (section 9.23 finding 6). */
  const [sending, setSending] = useState<string | null>(null);
  const [error, setError] = useState<string | null>(null);
  const [answer, setAnswer] = useState<Answer | null>(null);
  const [pasted, setPasted] = useState('');
  const abort = useRef<AbortController | null>(null);
  useEffect(() => { if (reset > 0) { setRequest(null); setAnswer(null); setError(null); setPasted(''); } }, [reset]);
  /** VS Code's models: listed when the box first shows, never at the page's start (P46), then kept current by the host. */
  const [vsModels, setVsModels] = useState<HostModel[] | null>(null);
  useEffect(() => {
    if (!hostAsk) return;
    let live = true;
    const off = hostAsk.onModels((list) => { if (live) setVsModels(list); });
    void hostAsk.models().then((list) => { if (live) setVsModels(list); }, () => { if (live) setVsModels([]); });
    return () => { live = false; off(); };
  }, [hostAsk]);
  const keyStored = useHostKeyStored(hostAsk);
  /** In a browser, a key in Options: without one, the request is copied to a chat and the answer pasted back (any plan, even a free one). */
  const browserKey = !hostAsk && askOption(OPTION_ASK_KEY) !== '';
  /** The model chosen: in VS Code as P46 has it (kept, else the first Claude, else the first); in a browser, kept or the first. */
  const choice = askChoice(hostAsk !== undefined, kept, vsModels ?? [], hostAsk ? keyStored : browserKey);
  const { model, route } = choice;

  /** An answer, from the provider or pasted from a chat: parsed into its form, a script handed to the console unrun. */
  const receive = useCallback(async (raw: AskAnswer): Promise<void> => {
    const form = parseAnswer(raw.text);
    const next: Answer = { raw, form };
    if (form.kind === 'wml') {
      try { next.elements = await parseElements(form.xml, 'block'); } catch (e) { next.parseError = (e as Error).message; }
    } else if (form.kind === 'script') {
      onScript(form.code);
    }
    setAnswer(next);
    setRequest(null);
    setPasted('');
  }, [onScript]);

  const send = useCallback(async (pending: Pending) => {
    const { request: req, route: via } = pending;
    const apiKey = askOption(OPTION_ASK_KEY);
    if (via === 'chat' || (via === 'browser' && apiKey === '') || ((via === 'vscode' || via === 'key') && !hostAsk)) { setRequest(pending); return; }
    setSending(pending.label);
    setError(null);
    setAnswer(null);
    abort.current = new AbortController();
    try {
      let raw: AskAnswer;
      if (via === 'vscode' || via === 'key') raw = await hostAsk!.send(via, req, abort.current.signal);
      else {
        const baseURL = askOption(OPTION_ASK_BASE_URL);
        raw = await anthropicProvider.send(req, { apiKey, ...(baseURL ? { baseURL } : {}), signal: abort.current.signal });
      }
      await receive(raw);
      onStatus(raw.refusal ? t('ask.declined') : raw.usage ? t('ask.answeredUsage', { model: raw.model, input: raw.usage.input, output: raw.usage.output }) : t('ask.answered', { model: raw.model }));
    } catch (e) {
      if (wasCancelled(e)) onStatus(t('ask.cancelled'));
      else {
        setError(describeError(e));
        // A model that is gone: the list is asked for again, so that the box offers what VS Code offers now.
        if (e instanceof AskFailure && e.failure.kind === 'notFound') void hostAsk?.models(true).catch(() => undefined);
      }
    } finally {
      setSending(null);
      abort.current = null;
    }
  }, [receive, onStatus, hostAsk]);

  /** The request as one message on the clipboard, for a chat. */
  const copyPrompt = useCallback(async (req: AskRequest) => {
    try {
      await navigator.clipboard.writeText(chatPrompt(req));
      onStatus(t('ask.chat.copied'));
    } catch (e) {
      onStatus(t('ask.chat.copyFailed', { error: (e as Error).message }));
    }
  }, [onStatus]);

  const usePasted = useCallback(async () => {
    if (pasted.trim() === '') return;
    setError(null);
    await receive({ text: pasted, model: '', stopReason: null });
    onStatus(t('ask.chat.pasted'));
  }, [pasted, receive, onStatus]);

  const ask = useCallback(async () => {
    if (prompt.trim() === '') return;
    setError(null);
    // Before VS Code's first listing has come, a kept VS Code model is not yet offered: the request waits for the listing,
    // and never goes by another route meanwhile (section 9.23 finding 5).
    const chosen = hostAsk && vsModels === null ? askChoice(true, kept, await hostAsk.models().catch(() => []), keyStored) : choice;
    const ctx = await buildContext(session, editor);
    // A VS Code model is named to the extension by its key, `<vendor>/<id>` (section 9.20 finding 1).
    const requestModel = chosen.vsModel ? modelKey(chosen.vsModel) : chosen.model ?? ASK_MODELS[0]!.id;
    const pending: Pending = { request: buildRequest(requestModel, prompt, ctx), route: chosen.route, label: chosen.label };
    if (sendAtOnce && chosen.route !== 'chat') await send(pending);
    else setRequest(pending);
  }, [prompt, choice, hostAsk, vsModels, kept, keyStored, sendAtOnce, session, editor, send]);

  /** Insert the answer after the last selected block (a text answer at the caret). */
  const insert = useCallback(async () => {
    if (!answer) return;
    const { form } = answer;
    if (form.kind === 'text') {
      const { to } = editor.state.selection;
      editor.caretAt(to);
      editor.run('insertText', { text: form.text });
      onStatus(t('ask.inserted'));
      return;
    }
    if (form.kind !== 'wml' || !answer.elements) return;
    const spec = selectionSpec(editor.state);
    const xml = form.xml;
    const result = await session.runOnTree(async ({ body }) => {
      const last = selectionBinding(body, spec).paragraphs.at(-1);
      if (last) await body.insertXml(xml, 'After', last); else await body.insertXml(xml, 'End');
    });
    if (result.error) setError((result.error as Error).message);
    else onStatus(t('ask.insertedAfter'));
  }, [answer, editor, session, onStatus]);

  /** Replace the selected blocks with the answer (a text answer replaces the selected text). */
  const replace = useCallback(async () => {
    if (!answer) return;
    const { form } = answer;
    if (form.kind === 'text') {
      editor.run('insertText', { text: form.text });
      onStatus(t('ask.replacedText'));
      return;
    }
    if (form.kind !== 'wml' || !answer.elements) return;
    const blocks = blocksIn(editor.state);
    const elements = answer.elements;
    const result = await session.runOnTree(({ body }) => {
      body.content.splice(blocks.from, blocks.to - blocks.from + 1, ...(elements as never[]));
    });
    if (result.error) setError((result.error as Error).message);
    else onStatus(t('ask.replacedBlocks', { count: blocks.nodes.length }));
  }, [answer, editor, session, onStatus]);

  const copy = useCallback(async (what: 'text' | 'xml' | 'script') => {
    if (!answer) return;
    const { form } = answer;
    let text = '';
    if (what === 'xml' && form.kind === 'wml') text = form.xml;
    else if (what === 'script' && form.kind === 'wml' && answer.elements) text = await toApiScript(answer.elements).catch((e: Error) => `// ${e.message}`);
    else if (form.kind === 'text') text = form.text;
    else if (form.kind === 'script') text = form.code;
    else if (form.kind === 'explanation') text = form.text;
    else if (answer.elements) text = answer.elements.map((el) => textOfElement(el)).join('\n');
    await navigator.clipboard?.writeText(text);
    onStatus(t('ask.copied'));
  }, [answer, onStatus]);

  const pickModel = (id: string): void => { setKept(id); setAskOption(OPTION_ASK_MODEL, id); };
  const chooseSendAtOnce = (on: boolean): void => { setSendAtOnce(on); setAskOption(OPTION_ASK_SEND_AT_ONCE, on ? 'true' : ''); };

  return (
    <Fold id="ask" level={3} className="ask-box" label={t('ask.title')} testId="ask-box" title={t('ask.title')}>
      <p className="meta">{hostAsk ? t('ask.note.vscode') : t('ask.note')}</p>
      <textarea data-testid="ask-prompt" aria-label={t('ask.question')} rows={3} value={prompt} onChange={(e) => setPrompt(e.target.value)}
        // With nothing offered there is no list to open: the models are asked for again as the person starts a question.
        onFocus={() => { if (hostAsk && model === undefined) void hostAsk.models(true).catch(() => undefined); }}
        placeholder={t('ask.question.placeholder')}
        onKeyDown={(e) => { if (e.key === 'Enter' && (e.ctrlKey || e.metaKey)) { e.preventDefault(); void ask(); } }} />
      <div className="pane-actions">
        {hostAsk ? (model !== undefined && (
          // VS Code's models first, then the Anthropic API's with the key (P46); asked for again when the list is opened, since
          // VS Code does not say when a provider registers (section 9.21).
          <select data-testid="ask-model" aria-label={t('ask.model')} value={model} onChange={(e) => pickModel(e.target.value)}
            onFocus={() => { void hostAsk.models(true).catch(() => undefined); }}>
            {(vsModels ?? []).length > 0 && (
              <optgroup label={t('ask.group.vscode')} data-testid="ask-group-vscode">
                {(vsModels ?? []).map((m) => <option key={modelKey(m)} value={`${VSCODE_MODEL}${modelKey(m)}`}>{t('ask.model.vscode', { name: m.name, vendor: m.vendor })}</option>)}
              </optgroup>
            )}
            {keyStored && (
              <optgroup label={t('ask.group.key')} data-testid="ask-group-key">
                {ASK_MODELS.map((m) => <option key={m.id} value={m.id}>{m.label}</option>)}
              </optgroup>
            )}
          </select>
        )) : (
          <select data-testid="ask-model" aria-label={t('ask.model')} value={model} onChange={(e) => pickModel(e.target.value)}>
            {ASK_MODELS.map((m, i) => <option key={m.id} value={m.id}>{i === 0 ? t('ask.model.default', { model: m.label }) : m.label}</option>)}
          </select>
        )}
        {(!hostAsk || route !== 'chat') && <label><input type="checkbox" data-testid="ask-send-at-once" checked={sendAtOnce} onChange={(e) => chooseSendAtOnce(e.target.checked)} /> {t('ask.sendAtOnce')}</label>}
        <button type="button" data-testid="ask-ask" onClick={() => void ask()} disabled={sending !== null || prompt.trim() === ''}>{t('ask.ask')}</button>
        {sending !== null && <button type="button" onClick={() => abort.current?.abort()}>{t('ask.cancel')}</button>}
      </div>
      {request && sending === null && (
        <div className="ask-request" data-testid="ask-request" data-route={request.route}>
          {request.route === 'vscode'
            // No system role in VS Code's API (section 4.6): the request goes as the chat route's one message.
            ? <p data-testid="ask-one-message">{tNodes('ask.request.vscode', { model: <b>{request.label}</b> })}</p>
            : request.route !== 'chat'
              ? <p>{tNodes('ask.request', { provider: anthropicProvider.name, model: <b>{request.request.model}</b> })}</p>
              : <p data-testid="ask-no-key">{hostAsk ? t('ask.noKey.vscode', { command: setKeyCommand() }) : t('ask.noKey')}</p>}
          <details>
            <summary>{t('ask.request.system', { count: formatNumber(request.request.system.length) })}</summary>
            <pre>{request.request.system}</pre>
          </details>
          <details open>
            <summary>{t('ask.request.message', { count: formatNumber(request.request.prompt.length) })}</summary>
            <pre>{request.request.prompt}</pre>
          </details>
          <div className="dialog-buttons">
            <button type="button" onClick={() => setRequest(null)}>{t('ask.cancel')}</button>
            <button type="button" data-testid="ask-copy-prompt" title={t('ask.chat.copy.title')} onClick={() => void copyPrompt(request.request)}>{t('ask.chat.copy')}</button>
            {request.route !== 'chat' && <button type="button" data-testid="ask-send" onClick={() => void send(request)}>{t('ask.send')}</button>}
          </div>
          <div className="ask-paste" data-testid="ask-paste-box">
            <p className="meta">{tNodes('ask.chat.note', { link: <a href="https://claude.ai/new" target="_blank" rel="noopener noreferrer">claude.ai</a> })}</p>
            <textarea data-testid="ask-paste" aria-label={t('ask.chat.paste')} rows={4} value={pasted} placeholder={t('ask.chat.paste.placeholder')} onChange={(e) => setPasted(e.target.value)} />
            <div className="pane-actions">
              <button type="button" data-testid="ask-use-pasted" disabled={pasted.trim() === ''} onClick={() => void usePasted()}>{t('ask.chat.use')}</button>
            </div>
          </div>
        </div>
      )}
      {sending !== null && <p className="meta" data-testid="ask-sending">{t('ask.sending', { model: sending })}</p>}
      {error && <p className="error" role="alert" data-testid="ask-error">{error}</p>}
      {answer && (
        <div className="ask-answer" data-testid="ask-answer" data-form={answer.raw.refusal ? 'refusal' : answer.form.kind}>
          {answer.raw.refusal ? (
            <p className="error">{t('ask.refusal', { category: answer.raw.refusal.category ? t('ask.refusal.category', { category: answer.raw.refusal.category }) : '', explanation: answer.raw.refusal.explanation ? t('ask.refusal.explanation', { explanation: answer.raw.refusal.explanation }) : '' })}</p>
          ) : answer.form.kind === 'explanation' ? (
            <pre className="ask-explanation" data-testid="ask-explanation">{answer.form.text}</pre>
          ) : (
            <>
              {answer.form.note !== '' && <p className="meta">{answer.form.note}</p>}
              {answer.form.kind === 'wml' && (answer.elements
                ? <AskPreview session={session} elements={answer.elements} />
                : <p className="error" role="alert">{t('ask.parseError', { error: answer.parseError ?? '' })}</p>)}
              {answer.form.kind === 'text' && <pre className="ask-text" data-testid="ask-text">{answer.form.text}</pre>}
              {answer.form.kind === 'script' && <p data-testid="ask-script-note">{t('ask.scriptNote')}</p>}
              <div className="pane-actions">
                {(answer.form.kind === 'text' || (answer.form.kind === 'wml' && answer.elements)) && (
                  <>
                    <button type="button" data-testid="ask-insert" onClick={() => void insert()}>{t('ask.insert')}</button>
                    <button type="button" data-testid="ask-replace" onClick={() => void replace()}>{t('ask.replace')}</button>
                  </>
                )}
                <button type="button" onClick={() => void copy('text')}>{answer.form.kind === 'wml' ? t('ask.copyText') : t('ask.copy')}</button>
                {answer.form.kind === 'wml' && <button type="button" onClick={() => void copy('xml')}>{t('ask.copyXml')}</button>}
                {answer.form.kind === 'wml' && answer.elements && <button type="button" onClick={() => void copy('script')}>{t('ask.copyScript')}</button>}
              </div>
            </>
          )}
        </div>
      )}
    </Fold>
  );
}

/** The text of a block element: its `w:t` values in order, a kept `mc:AlternateContent` by its selected branch alone. */
function textOfElement(el: Element): string {
  const out: string[] = [];
  const walk = (v: unknown): void => {
    if (v === null || typeof v !== 'object') return;
    if (Array.isArray(v)) { v.forEach(walk); return; }
    const o = v as { name?: { localPart?: string }; value?: unknown };
    if (o.name?.localPart === 't' && typeof (o.value as { value?: unknown })?.value === 'string') { out.push((o.value as { value: string }).value); return; }
    if (o.name?.localPart === 'AlternateContent') { walk(selectedBranchOf(o.value as Parameters<typeof selectedBranchOf>[0])); return; }
    for (const [k, x] of Object.entries(o)) if (k !== 'PARENT') walk(x);
  };
  walk(el);
  return out.join('');
}

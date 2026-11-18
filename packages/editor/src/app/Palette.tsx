// SPDX-License-Identifier: GPL-3.0-or-later
//
// The command palette (ED-001 section 3.6; ED-002 section 5.4): every registered command by
// label, with its key, in every perspective. A command with arguments asks for them from
// `spec.args`. A command can also be inserted in the console as the line that would run it
// (`editor.run(name, args)`, Shift+Enter or the row's button), inserted only, never run: the
// way to learn the names and the shapes, and to build a script from what the palette does.
import { useEffect, useMemo, useRef, useState } from 'react';
import { listCommands, WORD_KEYMAP, type CommandSpec } from '@docx4j/editor-model';
import type { EditorSurface } from '../view/surface.js';
import { useT, useLanguage } from './i18n.js';
import { commandLabel, argumentDescription } from './labels.js';

const keyOf = new Map<string, string>();
for (const b of WORD_KEYMAP) if (!keyOf.has(b.command)) keyOf.set(b.command, b.key.replace('Mod-', 'Ctrl+').replace('Shift-', 'Shift+').replace('Alt-', 'Alt+'));

/** 'true' | 'false' | a number | 'null' | 'toggle' | anything else as text. */
function parseArg(text: string): unknown {
  const t = text.trim();
  if (t === 'true') return true;
  if (t === 'false') return false;
  if (t === 'null' || t === '') return null;
  if (/^-?\d+(\.\d+)?$/.test(t)) return Number(t);
  return t;
}

/**
 * The console line for a command: its portable content-API form when it has one and every
 * argument is given (ED-003 section 7.2), else `editor.run('name', { key: value })`, an argument
 * not given left as `undefined` with its description.
 */
export function consoleLine(spec: CommandSpec<unknown>, args: Record<string, string>): string {
  const keys = Object.keys(spec.args ?? {});
  if (spec.api && keys.every((k) => (args[k] ?? '').trim() !== '')) {
    const line = spec.api(keys.length === 0 ? undefined : Object.fromEntries(keys.map((k) => [k, parseArg(args[k]!)])));
    if (line !== undefined) return line;
  }
  if (keys.length === 0) return `editor.run(${JSON.stringify(spec.name)});`;
  const fields = keys.map((k) => {
    const text = args[k] ?? '';
    if (text.trim() === '') return `${k}: undefined /* ${(spec.args![k] ?? '').replace(/\*\//g, '* /')} */`;
    return `${k}: ${JSON.stringify(parseArg(text))}`;
  });
  return `editor.run(${JSON.stringify(spec.name)}, { ${fields.join(', ')} });`;
}

/** An action of the application's the palette lists (ED-005 section 6.4: the views of the tracked changes): data, not a registry entry. */
export interface PaletteAction {
  id: string;
  label: string;
  run: () => void;
}

/** A macro as the palette lists it (ED-003 section 7.4): data, not a registry entry. */
export interface PaletteMacro {
  name: string;
  home: 'browser' | 'document';
  run: () => void;
}

export function Palette({ editor, onClose, onInsert, macros = [], actions = [] }: { editor: EditorSurface; onClose: () => void; onInsert?: (script: string) => void; macros?: readonly PaletteMacro[]; actions?: readonly PaletteAction[] }) {
  const t = useT();
  const lang = useLanguage();
  const [query, setQuery] = useState('');
  const [index, setIndex] = useState(0);
  const [pending, setPending] = useState<CommandSpec<unknown> | null>(null);
  const [args, setArgs] = useState<Record<string, string>>({});
  const input = useRef<HTMLInputElement>(null);   // the query box, or the first argument field (a text input or a select)
  // By label in the current language; a query matches that label, the English one and the name.
  const commands = useMemo(() => listCommands().map((spec) => ({ spec, label: commandLabel(spec) })).sort((a, b) => a.label.localeCompare(b.label, lang)).map((c) => c.spec), [lang]);
  const matches = useMemo(() => {
    const q = query.toLocaleLowerCase(lang);
    return commands.filter((c) => commandLabel(c).toLocaleLowerCase(lang).includes(q) || c.label.toLowerCase().includes(q) || c.name.toLowerCase().includes(q));
  }, [commands, query, lang]);
  const macroMatches = useMemo(() => {
    const q = query.toLowerCase();
    return macros.filter((m) => t('palette.macro', { name: m.name }).toLowerCase().includes(q) || `macro: ${m.name}`.toLowerCase().includes(q));
  }, [macros, query, t]);
  const actionMatches = useMemo(() => {
    const q = query.toLocaleLowerCase(lang);
    return actions.filter((a) => a.label.toLocaleLowerCase(lang).includes(q));
  }, [actions, query, lang]);
  const total = matches.length + macroMatches.length + actionMatches.length;
  useEffect(() => { input.current?.focus(); }, [pending]);
  /** Insert the command's line in the console rather than running it: the arguments given so far, the rest as placeholders. */
  const insert = (spec: CommandSpec<unknown>): void => {
    if (!onInsert) return;
    onClose();
    onInsert(consoleLine(spec, pending === spec ? args : {}));
  };
  const run = (spec: CommandSpec<unknown>): void => {
    if (spec.args && Object.keys(spec.args).length > 0 && !pending) { setPending(spec); setArgs({}); return; }
    const a = spec.args ? Object.fromEntries(Object.keys(spec.args).map((k) => [k, parseArg(args[k] ?? '')])) : undefined;
    onClose();
    editor.run(spec.name, a);
    editor.focus();
  };
  return (
    <div className="dialog-backdrop" onMouseDown={(e) => { if (e.target === e.currentTarget) onClose(); }}>
      <div className="palette" role="dialog" aria-label={t('palette.label')} data-testid="palette"
        onKeyDown={(e) => {
          if (e.key === 'Escape') { onClose(); editor.focus(); }
          else if (!pending && e.key === 'ArrowDown') { e.preventDefault(); setIndex((i) => Math.min(i + 1, total - 1)); }
          else if (!pending && e.key === 'ArrowUp') { e.preventDefault(); setIndex((i) => Math.max(i - 1, 0)); }
          else if (e.key === 'Enter' && e.shiftKey && onInsert) { e.preventDefault(); const spec = pending ?? matches[index]; if (spec) insert(spec); }
          else if (e.key === 'Enter') {
            e.preventDefault();
            const spec = pending ?? matches[index];
            if (spec) run(spec);
            else {
              const m = macroMatches[index - matches.length];
              const a = actionMatches[index - matches.length - macroMatches.length];
              if (m) { onClose(); m.run(); } else if (a) { onClose(); a.run(); }
            }
          }
        }}>
        {!pending ? (
          <>
            <input ref={input} type="text" placeholder={t('palette.query')} aria-label={t('palette.query.label')} value={query} onChange={(e) => { setQuery(e.target.value); setIndex(0); }} />
            <ul role="listbox" aria-label={t('palette.commands')}>
              {matches.map((c, i) => (
                <li key={c.name} role="option" aria-selected={i === index} className={i === index ? 'selected' : ''} data-command={c.name}
                  onMouseEnter={() => setIndex(i)} onMouseDown={(e) => e.preventDefault()} onClick={() => run(c)}>
                  <span className="label">{commandLabel(c)}</span>
                  <span className="name">{c.name}</span>
                  {keyOf.has(c.name) ? <kbd>{keyOf.get(c.name)}</kbd> : <span />}
                  <span className={`badge ${c.api ? 'api' : 'editor-only'}`} title={c.api ? t('palette.badge.api.title') : t('palette.badge.editor.title')}>{c.api ? t('palette.badge.api') : t('palette.badge.editor')}</span>
                  {onInsert && <button type="button" className="link insert" title={t('palette.insert.title')} aria-label={t('palette.insert.label', { command: commandLabel(c) })} data-testid="palette-insert" onClick={(e) => { e.stopPropagation(); insert(c); }}>{t('palette.insert')}</button>}
                </li>
              ))}
              {macroMatches.map((m, k) => {
                const i = matches.length + k;
                return (
                  <li key={`macro:${m.home}:${m.name}`} role="option" aria-selected={i === index} className={i === index ? 'selected' : ''} data-macro={m.name}
                    onMouseEnter={() => setIndex(i)} onMouseDown={(e) => e.preventDefault()} onClick={() => { onClose(); m.run(); }}>
                    <span className="label">{t('palette.macro', { name: m.name })}</span>
                    <span className="name">{t('palette.macro.undo')}</span>
                    <span />
                    <span className={`badge macro-${m.home}`} title={m.home === 'document' ? t('macros.home.document.title') : t('macros.home.browser.title')}>{t(`macros.home.${m.home}`)}</span>
                    <span />
                  </li>
                );
              })}
              {actionMatches.map((a, k) => {
                const i = matches.length + macroMatches.length + k;
                return (
                  <li key={a.id} role="option" aria-selected={i === index} className={i === index ? 'selected' : ''} data-action={a.id}
                    onMouseEnter={() => setIndex(i)} onMouseDown={(e) => e.preventDefault()} onClick={() => { onClose(); a.run(); }}>
                    <span className="label">{a.label}</span>
                    <span /><span /><span /><span />
                  </li>
                );
              })}
            </ul>
          </>
        ) : (
          <form onSubmit={(e) => { e.preventDefault(); run(pending); }}>
            <div className="label">{commandLabel(pending)}</div>
            {Object.entries(pending.args ?? {}).map(([name, description], i) => {
              const choices = pending.choices?.[name];
              return (
                <div key={name} className="palette-field">
                  <label>
                    <span>{name}</span>
                    {choices
                      ? (
                        <select ref={i === 0 ? (input as unknown as React.RefObject<HTMLSelectElement>) : undefined} value={args[name] ?? ''} onChange={(e) => setArgs({ ...args, [name]: e.target.value })}>
                          <option value="">{t('palette.notGiven')}</option>
                          {choices.map((c) => <option key={c} value={c}>{c}</option>)}
                        </select>
                      )
                      : <input ref={i === 0 ? input : undefined} type="text" value={args[name] ?? ''} onChange={(e) => setArgs({ ...args, [name]: e.target.value })} />}
                  </label>
                  <p className="meta">{argumentDescription(pending, name, description)}</p>
                </div>
              );
            })}
            <div className="dialog-buttons">
              <button type="button" onClick={() => setPending(null)}>{t('palette.back')}</button>
              {onInsert && <button type="button" title={t('palette.insertArgs.title')} data-testid="palette-insert-args" onClick={() => insert(pending)}>{t('palette.insertArgs')}</button>}
              <button type="submit">{t('palette.run')}</button>
            </div>
          </form>
        )}
      </div>
    </div>
  );
}

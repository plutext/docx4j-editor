// SPDX-License-Identifier: GPL-3.0-or-later
//
// The options every editor of the extension shares (P12 revised, P19): kept in `globalState`, written into each page as
// boot data, and a page's change passed on to every other open page.
import type * as vscode from 'vscode';
import { isSharedOption, type OptionsMessage } from './shared.js';

const KEY = 'docx4j.options';

export class OptionsStore {
  constructor(private readonly memento: vscode.Memento) {}

  all(): Record<string, string> {
    const kept = this.memento.get<Record<string, unknown>>(KEY, {});
    const out: Record<string, string> = {};
    for (const [key, value] of Object.entries(kept)) if (isSharedOption(key) && typeof value === 'string') out[key] = value;
    return out;
  }

  /** A page's change, kept; only shared keys, as the page posts them. */
  async apply(message: OptionsMessage): Promise<OptionsMessage | undefined> {
    const next = this.all();
    const set: Record<string, string> = {};
    const remove: string[] = [];
    for (const [key, value] of Object.entries(message.set ?? {})) if (isSharedOption(key) && typeof value === 'string') { next[key] = value; set[key] = value; }
    for (const key of message.remove ?? []) if (isSharedOption(key) && key in next) { delete next[key]; remove.push(key); }
    if (Object.keys(set).length === 0 && remove.length === 0) return undefined;
    await this.memento.update(KEY, next);
    return { type: 'options', ...(Object.keys(set).length > 0 ? { set } : {}), ...(remove.length > 0 ? { remove } : {}) };
  }
}

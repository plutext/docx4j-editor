// SPDX-License-Identifier: GPL-3.0-or-later
//
// The Ask box's model in VS Code (ED-014 section 4.6, P46): the kept choice while offered, else the first Claude, else the
// first listed; a VS Code model kept as `vscode:<vendor>/<id>`.
import { describe, expect, test } from 'vitest';
import { askChoice, chooseModel, vscodeModelOf, ASK_MODELS } from '../src/app/ask.js';
import type { HostModelData } from '../src/app/host-protocol.js';

const m = (vendor: string, id: string, family = id): HostModelData => ({ vendor, id, family, name: `${vendor} ${id}`, version: '1', maxInputTokens: 1000 });
const gpt = m('copilot', 'gpt-x');
const claudeCopilot = m('copilot', 'claude-sonnet-x');
const claudeKey = m('anthropic', 'claude-sonnet-x');
const fam = m('other', 'model-7', 'Claude-like');

describe('chooseModel', () => {
  test('the kept choice while it is offered, a VS Code model or, with a key, an Anthropic one', () => {
    expect(chooseModel('vscode:copilot/gpt-x', [gpt, claudeCopilot], false)).toBe('vscode:copilot/gpt-x');
    expect(chooseModel('claude-opus-5', [gpt], true)).toBe('claude-opus-5');
  });

  test('else the first VS Code model whose family or id names Claude', () => {
    expect(chooseModel('', [gpt, claudeCopilot, claudeKey], true)).toBe('vscode:copilot/claude-sonnet-x');
    expect(chooseModel('vscode:gone/x', [gpt, fam], false)).toBe('vscode:other/model-7');
    // An Anthropic model kept with no key stored is not offered.
    expect(chooseModel('claude-opus-5', [gpt, claudeKey], false)).toBe('vscode:anthropic/claude-sonnet-x');
  });

  test('else the first listed: a VS Code model, or the key\'s first; none with neither', () => {
    expect(chooseModel('', [gpt], true)).toBe('vscode:copilot/gpt-x');
    expect(chooseModel('', [], true)).toBe(ASK_MODELS[0]!.id);
    expect(chooseModel('claude-opus-5', [], false)).toBeUndefined();
  });

  test('two vendors\' models of one id are two choices', () => {
    expect(vscodeModelOf('vscode:anthropic/claude-sonnet-x', [claudeCopilot, claudeKey])).toBe(claudeKey);
    expect(vscodeModelOf('vscode:copilot/claude-sonnet-x', [claudeCopilot, claudeKey])).toBe(claudeCopilot);
    expect(vscodeModelOf('claude-sonnet-x', [claudeCopilot])).toBeUndefined();
    expect(vscodeModelOf('vscode:copilot/none', [claudeCopilot])).toBeUndefined();
  });
});

describe('askChoice', () => {
  test('in VS Code: a VS Code model by its own route and name; an Anthropic model by the key while one is stored; else the chat', () => {
    expect(askChoice(true, 'vscode:copilot/gpt-x', [gpt], false)).toEqual({ model: 'vscode:copilot/gpt-x', vsModel: gpt, route: 'vscode', label: 'copilot gpt-x' });
    expect(askChoice(true, 'claude-opus-5', [gpt], true)).toEqual({ model: 'claude-opus-5', vsModel: undefined, route: 'key', label: 'Claude Opus 5' });
    expect(askChoice(true, '', [], false)).toEqual({ model: undefined, vsModel: undefined, route: 'chat', label: '' });
  });

  test('before the listing, a kept VS Code model is not offered, and the choice goes by another route: why a request waits for the listing (section 9.23 finding 5)', () => {
    expect(askChoice(true, 'vscode:copilot/gpt-x', [], true).route).toBe('key');
    expect(askChoice(true, 'vscode:copilot/gpt-x', [gpt], true).route).toBe('vscode');
  });

  test('in a browser: kept or the first, by the key in Options or the chat; VS Code\'s models unused', () => {
    expect(askChoice(false, '', [gpt], true)).toEqual({ model: ASK_MODELS[0]!.id, vsModel: undefined, route: 'browser', label: ASK_MODELS[0]!.label });
    expect(askChoice(false, 'claude-opus-5', [], false)).toEqual({ model: 'claude-opus-5', vsModel: undefined, route: 'chat', label: 'Claude Opus 5' });
  });
});

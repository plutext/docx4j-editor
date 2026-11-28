// SPDX-License-Identifier: GPL-3.0-or-later
//
// The suite's driver (ED-014 P31): VS Code's side of the suite in VS Code, never packaged. It runs commands, reads the tabs
// and the clipboard, and answers on a loopback port, `DOCX4J_DRIVER_PORT`, which the suite gives it. The page itself the
// suite reaches over the Chrome DevTools Protocol, as a person does.
const vscode = require('vscode');
const http = require('node:http');
const { join } = require('node:path');
const { pathToFileURL } = require('node:url');

// ---- A language model for the Ask box (ED-014 section 4.6, S6.4): vendor `docx4j-test`, declared in the manifest, its
// answers the browser suite's double's (`answerFor`) over the question after the chat prompt's rule. What it receives, and
// whether its token was cancelled, the spec reads with `lmReceived`. A question that says SLOW waits for its cancellation.
// The suite passes its path in the environment, since an installed copy of the driver is not beside the fixtures (ED-014
// section 9.24 item 4); a development copy finds it by its own path too.
const DOUBLE = pathToFileURL(process.env.DOCX4J_DRIVER_DOUBLE || join(__dirname, '..', '..', '..', '..', 'editor-model', 'test', 'fixtures', 'build', 'ask-double.mjs')).href;
const RULE = '\n\n---\n\n';
const lmReceived = [];
let lmRegistration;
const lmModels = [
  { id: 'claude-test', name: 'Test Claude', family: 'claude-test', version: '1', maxInputTokens: 200000, maxOutputTokens: 4096, capabilities: {} },
  { id: 'claude-tiny', name: 'Tiny Claude', family: 'claude-test', version: '1', maxInputTokens: 1000, maxOutputTokens: 4096, capabilities: {} },
];
const partText = (p) => (p && typeof p.value === 'string' ? p.value : '');
const lmProvider = {
  provideLanguageModelChatInformation: async () => lmModels,
  async provideLanguageModelChatResponse(model, messages, options, progress, token) {
    const entry = { model: model.id, messages: messages.map((m) => ({ role: m.role, text: m.content.map(partText).join('') })), modelOptions: options.modelOptions ?? null, cancelled: false };
    lmReceived.push(entry);
    token.onCancellationRequested(() => { entry.cancelled = true; });
    const text = entry.messages.map((m) => m.text).join('\n');
    const at = text.indexOf(RULE);
    const question = at >= 0 ? text.slice(at + RULE.length) : text;
    if (question.startsWith('SLOW')) {
      for (let i = 0; i < 300 && !token.isCancellationRequested; i++) await new Promise((r) => setTimeout(r, 100));
      if (token.isCancellationRequested) throw new vscode.CancellationError();
    }
    const { answerFor } = await import(DOUBLE);
    progress.report(new vscode.LanguageModelTextPart(answerFor(question) ?? 'The double declines.'));
  },
  provideTokenCount: async (_model, text) => Math.ceil((typeof text === 'string' ? text : text.content.map(partText).join('')).length / 4),
};
function lmRegister() {
  if (!lmRegistration && typeof vscode.lm?.registerLanguageModelChatProvider === 'function') lmRegistration = vscode.lm.registerLanguageModelChatProvider('docx4j-test', lmProvider);
  return !!lmRegistration;
}
function lmUnregister() { lmRegistration?.dispose(); lmRegistration = undefined; return true; }

/** A `{ $uri }` in the arguments is a file URI; `{ $parse }` a URI parsed from its string. */
function revive(v) {
  if (Array.isArray(v)) return v.map(revive);
  if (v && typeof v === 'object') {
    if (typeof v.$uri === 'string') return vscode.Uri.file(v.$uri);
    if (typeof v.$parse === 'string') return vscode.Uri.parse(v.$parse);
    return Object.fromEntries(Object.entries(v).map(([k, x]) => [k, revive(x)]));
  }
  return v;
}

function plain(v) {
  try { return JSON.parse(JSON.stringify(v ?? null)); } catch { return String(v); }
}

function tabs() {
  return vscode.window.tabGroups.all.flatMap((g) => g.tabs.map((t) => {
    const input = t.input;
    return {
      group: g.viewColumn,
      label: t.label,
      isDirty: t.isDirty,
      isActive: t.isActive,
      viewType: input && 'viewType' in input ? input.viewType : null,
      kind: input instanceof vscode.TabInputCustom ? 'custom' : input instanceof vscode.TabInputText ? 'text' : input ? input.constructor.name : null,
      uri: input && input.uri ? input.uri.toString() : null,
    };
  }));
}

const notes = [];

let context;

async function op(body) {
  switch (body.op) {
    case 'ping': return { version: vscode.version, trusted: vscode.workspace.isTrusted, language: vscode.env.language };
    case 'exec': return plain(await vscode.commands.executeCommand(body.command, ...revive(body.args ?? [])));
    case 'tabs': return tabs();
    // The manifest an extension reads as `packageJSON`, as VS Code translated it (ED-014 section 4.7).
    case 'manifest': return plain(vscode.extensions.getExtension(body.id)?.packageJSON ?? null);
    case 'closeTab': {
      const t = vscode.window.tabGroups.all.flatMap((g) => g.tabs).find((x) => x.label === body.label);
      if (!t) throw new Error(`no tab ${body.label}`);
      return vscode.window.tabGroups.close(t);
    }
    case 'clipboardRead': return vscode.env.clipboard.readText();
    case 'clipboardWrite': await vscode.env.clipboard.writeText(body.text); return true;
    case 'config': await vscode.workspace.getConfiguration(body.section).update(body.key, body.value, vscode.ConfigurationTarget.Global); return true;
    case 'notes': return notes.slice(body.since ?? 0);
    // A notification toast, for the keys that hide one (ED-014 section 4.4); not awaited, since it settles when dismissed.
    case 'notify': void vscode.window.showInformationMessage(body.text); return true;
    // A rename as the explorer makes one: through the working copies, so a dirty editor's backup follows it.
    case 'rename': {
      const edit = new vscode.WorkspaceEdit();
      edit.renameFile(vscode.Uri.file(body.from), vscode.Uri.file(body.to));
      return vscode.workspace.applyEdit(edit);
    }
    case 'storagePath': return { workspace: context.storageUri?.fsPath ?? null, global: context.globalStorageUri.fsPath };
    // The language model tools VS Code knows, MCP servers' among them (ED-014 section 9.17 item 4).
    case 'lmTools': return vscode.lm.tools.map((t) => ({ name: t.name, tags: t.tags }));
    // A tool called as VS Code's agent calls it, outside a chat: VS Code asks the person first, which the suite answers.
    case 'invokeTool': {
      const r = await vscode.lm.invokeTool(body.name, { input: body.input ?? {}, toolInvocationToken: undefined });
      return r.content.map((c) => c.value ?? '').join('');
    }
    // The Ask box's test model (S6.4): what it received, and its registration.
    case 'lmReceived': return lmReceived.slice(body.since ?? 0);
    case 'lmRegister': return lmRegister();
    case 'lmUnregister': return lmUnregister();
    default: throw new Error(`unknown op ${body.op}`);
  }
}

function activate(ctx) {
  context = ctx;
  const port = Number(process.env.DOCX4J_DRIVER_PORT || 47621);
  const server = http.createServer((req, res) => {
    let s = '';
    req.on('data', (c) => { s += c; });
    req.on('end', async () => {
      try {
        const out = await op(JSON.parse(s || '{}'));
        res.writeHead(200, { 'content-type': 'application/json' });
        res.end(JSON.stringify({ ok: true, out: out ?? null }));
      } catch (e) {
        res.writeHead(200, { 'content-type': 'application/json' });
        res.end(JSON.stringify({ ok: false, error: String((e && e.stack) || e) }));
      }
    });
  });
  server.listen(port, '127.0.0.1');
  context.subscriptions.push({ dispose: () => server.close() }, { dispose: lmUnregister });
  // Registered from the start where the spec asks for it (`lm` in the harness's launch), before any page lists.
  if (process.env.DOCX4J_DRIVER_LM === 'on') lmRegister();
  context.subscriptions.push(vscode.window.tabGroups.onDidChangeTabs((e) => {
    for (const t of e.opened) notes.push({ t: Date.now(), kind: 'opened', label: t.label });
    for (const t of e.closed) notes.push({ t: Date.now(), kind: 'closed', label: t.label });
  }));
}

module.exports = { activate, deactivate() {} };

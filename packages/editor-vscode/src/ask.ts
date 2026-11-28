// SPDX-License-Identifier: GPL-3.0-or-later
//
// The Ask box in VS Code (ED-014 section 4.6, P44 to P46): every request the page sends goes through here, to one of VS
// Code's models (`vscode.lm`) or to the Anthropic API with the person's key. The key is kept in SecretStorage, set and
// removed by two commands, and never posted to a page; its address is the setting `docx4j.ask.baseUrl`, at application
// scope, which a page cannot write. The rules with no `vscode` import are `models.ts`'s.
import * as vscode from 'vscode';
import { AskService } from './models.js';
import { sendAnthropic, type HostModelData } from './shared.js';

/** The key's name in SecretStorage: one for the person, across windows and workspaces, as the MCP token is (P39). */
export const ASK_KEY = 'docx4j.ask.key';
const SECTION = 'docx4j.ask';
const ADDRESS = 'baseUrl';

export const askWords = {
  justification: (): string => vscode.l10n.t('The Ask box sends your question, the selected text and its XML, and the document\'s outline and styles to the model you chose.'),
};

export class Ask implements vscode.Disposable {
  readonly service: AskService<vscode.LanguageModelChat, unknown>;
  private readonly disposables: vscode.Disposable[] = [];
  private readonly keyChanged = new vscode.EventEmitter<boolean>();
  /** The key was set or removed, in this window or another: whether one is stored now. */
  readonly onDidChangeKey = this.keyChanged.event;
  private readonly modelsChanged = new vscode.EventEmitter<HostModelData[]>();
  /** VS Code's models changed, once a page has listed them: the list as it is now. */
  readonly onDidChangeModels = this.modelsChanged.event;

  constructor(private readonly context: vscode.ExtensionContext) {
    // A fork may lack the API (P30): then no model is listed, and the key and the chat remain.
    const lm = (vscode as { lm?: Partial<typeof vscode.lm> }).lm;
    const select = lm?.selectChatModels;
    this.service = new AskService<vscode.LanguageModelChat>({
      select: () => (typeof select === 'function' ? select.call(lm) : Promise.resolve([])),
      userMessage: (text) => vscode.LanguageModelChatMessage.User(text),
      tokenFor: (signal) => {
        const source = new vscode.CancellationTokenSource();
        const cancel = (): void => source.cancel();
        if (signal.aborted) source.cancel(); else signal.addEventListener('abort', cancel, { once: true });
        return { token: source.token, dispose: () => { signal.removeEventListener('abort', cancel); source.dispose(); } };
      },
      justification: askWords.justification,
      key: () => context.secrets.get(ASK_KEY),
      address: () => this.address(),
      sendKey: (request, options) => sendAnthropic(request, options),
    });
    this.disposables.push(this.keyChanged, this.modelsChanged);
    // Listed again on a change only once a page has asked: listing activates every vendor's extension (P46).
    if (typeof lm?.onDidChangeChatModels === 'function') {
      this.disposables.push(lm.onDidChangeChatModels(() => {
        if (this.service.listed) void this.service.models().then((models) => this.modelsChanged.fire(models));
      }));
    }
    this.disposables.push(context.secrets.onDidChange((e) => { if (e.key === ASK_KEY) void this.tellKey(); }));
  }

  dispose(): void {
    for (const d of this.disposables) d.dispose();
  }

  /**
   * The key's address: the user's setting, read at application scope, so a workspace's value is never taken (S6.1). A value
   * that is not text, written by hand into settings.json, is Anthropic's (section 9.23).
   */
  address(): string {
    const value: unknown = vscode.workspace.getConfiguration(SECTION).get(ADDRESS, '');
    return typeof value === 'string' ? value.trim() : '';
  }

  /** Whether a key is stored. */
  async keyStored(): Promise<boolean> {
    return (await this.context.secrets.get(ASK_KEY)) !== undefined;
  }

  private async tellKey(): Promise<void> {
    this.keyChanged.fire(await this.keyStored());
  }

  /** "docx4j editor: Set the Anthropic API Key…": a password box; what is typed goes to SecretStorage alone. */
  async setKey(): Promise<void> {
    const key = await vscode.window.showInputBox({
      title: vscode.l10n.t('Set the Anthropic API Key'),
      prompt: vscode.l10n.t('VS Code keeps the key in its secret storage, for every window. The Ask box sends with it to the address in the setting docx4j.ask.baseUrl, Anthropic\'s own when that is empty.'),
      placeHolder: 'sk-ant-…',
      password: true,
      ignoreFocusOut: true,
      validateInput: (value) => (/\s/.test(value.trim()) ? vscode.l10n.t('A key has no spaces in it.') : undefined),
    });
    if (key === undefined || key.trim() === '') return;
    await this.context.secrets.store(ASK_KEY, key.trim());
    await this.tellKey();
    void vscode.window.showInformationMessage(vscode.l10n.t('The Anthropic API key is set. The Ask box lists the Anthropic models beside VS Code\'s.'));
  }

  /** "docx4j editor: Remove the Anthropic API Key". */
  async removeKey(): Promise<void> {
    if (!(await this.keyStored())) {
      void vscode.window.showInformationMessage(vscode.l10n.t('No Anthropic API key is set.'));
      return;
    }
    await this.context.secrets.delete(ASK_KEY);
    await this.tellKey();
    void vscode.window.showInformationMessage(vscode.l10n.t('The Anthropic API key is removed.'));
  }

  /** The Options dialog's buttons (P45): the commands, or the address's setting. */
  action(action: unknown): void {
    switch (action) {
      case 'set': void vscode.commands.executeCommand('docx4j.ask.setKey'); break;
      case 'remove': void vscode.commands.executeCommand('docx4j.ask.removeKey'); break;
      case 'settings': void vscode.commands.executeCommand('workbench.action.openSettings', `${SECTION}.${ADDRESS}`); break;
    }
  }
}

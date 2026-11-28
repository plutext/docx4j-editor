// SPDX-License-Identifier: GPL-3.0-or-later
//
// The extension (ED-014): the application as a custom editor for the four Word types, the Markdown launcher, the New Word
// Document command, the commands that reach the page: Commands and Paste as Plain Text (section 4.4), and the agent's
// companion with its MCP server, Copy the Agent's MCP Address and Join a Room… (section 4.5), and the Ask box's key
// commands (section 4.6). Activated by the custom editors, the commands and the MCP provider its manifest contributes
// (VS Code 1.74 on).
import * as vscode from 'vscode';
import { DocxEditorProvider, MarkdownLauncher, MARKDOWN_VIEW_TYPE, WORD_VIEW_TYPE } from './provider.js';
import { OptionsStore } from './options.js';
import { Agent } from './agent.js';
import { Ask } from './ask.js';
import { paletteNames } from './names.js';
import { PAGE_COMMANDS } from './shared.js';

let agent: Agent | undefined;

export async function activate(context: vscode.ExtensionContext): Promise<void> {
  const root = vscode.Uri.joinPath(context.extensionUri, 'app');
  const indexHtml = new TextDecoder().decode(await vscode.workspace.fs.readFile(vscode.Uri.joinPath(root, 'index.html')));
  // The provider tells the agent which relays its pages' rooms are on, for the companion's restart rule.
  let word: DocxEditorProvider | undefined;
  agent = new Agent(context, (relay) => word?.roomsOn(relay) ?? false);
  // The Ask box's requests (section 4.6): VS Code's models, and the person's Anthropic key in SecretStorage.
  const ask = new Ask(context);
  // The palette's names of the commands the page names, as VS Code translated the manifest (section 4.7, P49).
  const commands = paletteNames(context.extension.packageJSON, Object.keys(PAGE_COMMANDS));
  word = new DocxEditorProvider({ root, indexHtml, commands }, new OptionsStore(context.globalState), agent, ask);
  agent.registerMcp();
  const words = word;
  // One editor per document, kept alive while hidden: the session lives in the page (P7).
  const webviewOptions = { retainContextWhenHidden: true, enableFindWidget: false };
  context.subscriptions.push(
    agent,
    ask,
    vscode.window.registerCustomEditorProvider(WORD_VIEW_TYPE, words, { webviewOptions, supportsMultipleEditorsPerDocument: false }),
    vscode.window.registerCustomEditorProvider(MARKDOWN_VIEW_TYPE, new MarkdownLauncher(words), { webviewOptions, supportsMultipleEditorsPerDocument: false }),
    // An untitled document, its page size asked as New asks; its first Save is VS Code's Save As (P17).
    vscode.commands.registerCommand('docx4j.newDocument', () => words.openUntitled('Document1.docx', new Uint8Array(), words.folderOf(undefined))),
    // The application's palette (P33), on Ctrl+Alt+P while this editor is active; and Paste as Plain Text from VS Code's
    // menus (P34), whose key is the page's own Ctrl+Shift+V (ED-014 section 4.4).
    vscode.commands.registerCommand('docx4j.commands', () => words.showCommands()),
    vscode.commands.registerCommand('docx4j.pasteText', () => words.pasteText()),
    // The agent's MCP address for Claude Code and any client, and a room's link opened in an editor (section 4.5).
    vscode.commands.registerCommand('docx4j.copyAgentAddress', () => agent?.copyAddress()),
    vscode.commands.registerCommand('docx4j.join', () => words.joinCommand()),
    // The Ask box's Anthropic key, kept in SecretStorage and never posted to a page (P45).
    vscode.commands.registerCommand('docx4j.ask.setKey', () => ask.setKey()),
    vscode.commands.registerCommand('docx4j.ask.removeKey', () => ask.removeKey()),
    // A dirty document renamed keeps its edits (ED-014 section 9.9), and one in a room its room (P42).
    vscode.workspace.onWillRenameFiles((e) => e.waitUntil(words.beforeRename(e.files))),
  );
}

export function deactivate(): void {
  // The companion is asked to write its rooms and exit; nothing is killed (ED-014 section 4.5). The rest goes with the context.
  agent?.close();
  agent = undefined;
}

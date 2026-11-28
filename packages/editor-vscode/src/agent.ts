// SPDX-License-Identifier: GPL-3.0-or-later
//
// The agent's companion in VS Code (ED-014 section 4.5, P36 to P43): its process, started on demand under VS Code's own Node
// with its ports and rooms kept per workspace; the MCP server VS Code's agent mode is offered once the person has admitted
// the agent in the workspace; and the address command for Claude Code and any other client. The rules with no `vscode`
// import are `companion.ts`'s.
import * as vscode from 'vscode';
import { spawn } from 'node:child_process';
import { randomBytes } from 'node:crypto';
import { createServer } from 'node:net';
import { join } from 'node:path';
import {
  CompanionRunner, canRunCompanion, choosePorts, claudeCodeLine, companionArgs, companionEnv, isPorts, mcpAddressOf,
  isWebviewOrigin, offered, relayAddressOf, roomsFolder, withOrigin, type CanRun, type ChildLike, type CompanionPorts, type Launch, type Running,
} from './companion.js';

const PORTS = 'docx4j.agent.ports';
const ADMITTED = 'docx4j.agent.admitted';
const ORIGINS = 'docx4j.agent.origins';
const TOKEN = 'docx4j.agent.mcpToken';
/** VS Code finds the server to start by its label (`$resolveMcpLaunch`), so it stays fixed (section 4.5). */
const SERVER_LABEL = 'docx4j editor';
export const MCP_PROVIDER = 'docx4j.companion';

/** Whether a port on loopback is free now: bound, then let go. */
function isFree(port: number): Promise<boolean> {
  return new Promise((resolve) => {
    const server = createServer();
    server.once('error', () => resolve(false));
    server.listen(port, '127.0.0.1', () => server.close(() => resolve(true)));
  });
}

export const agentWords = {
  cannotRun: (why: Exclude<CanRun, { ok: true }>['why']): string => {
    switch (why) {
      case 'untrusted': return vscode.l10n.t('The agent\'s companion is not started in Restricted Mode: trust the workspace to use it.');
      case 'remote': return vscode.l10n.t('The agent\'s companion runs only in a local window, not a remote one.');
      case 'virtual': return vscode.l10n.t('The agent\'s companion runs only in a workspace on this computer\'s disk.');
      case 'node': return vscode.l10n.t('The agent\'s companion needs Node 22 or later, and this VS Code runs an older one.');
    }
  },
  gaveUp: (): string => vscode.l10n.t('The agent\'s companion stopped five times in a minute and is left stopped: reload the window to start it again.'),
  reopen: (): string => vscode.l10n.t('The agent\'s companion moved to another port: close this document and open it again to share it or invite the agent.'),
  failed: (reason: string): string => vscode.l10n.t('The agent\'s companion did not start: {0}', reason),
};

export class Agent implements vscode.Disposable {
  private readonly runner: CompanionRunner;
  private readonly output: vscode.OutputChannel;
  private readonly changed = new vscode.EventEmitter<void>();
  private readonly disposables: vscode.Disposable[] = [];
  /** Told when the companion listens: every page then gets its address and its relay's. */
  private readonly startedEmitter = new vscode.EventEmitter<Running>();
  readonly onDidStart = this.startedEmitter.event;

  constructor(private readonly context: vscode.ExtensionContext, private readonly roomsOn: (relay: string) => boolean) {
    this.output = vscode.window.createOutputChannel('docx4j editor: Agent');
    this.runner = new CompanionRunner({
      spawn: (launch) => this.spawnChild(launch),
      launch: () => this.launch(),
      log: (line) => this.output.appendLine(line),
      roomsOn: (relay) => this.roomsOn(relay),
      started: (running) => this.startedEmitter.fire(running),
    });
    this.disposables.push(
      this.output, this.changed, this.startedEmitter,
      vscode.workspace.onDidChangeWorkspaceFolders(() => this.runner.post({ type: 'roots', roots: this.roots() })),
      vscode.workspace.onDidGrantWorkspaceTrust(() => this.changed.fire()),
    );
  }

  dispose(): void {
    for (const d of this.disposables) d.dispose();
  }

  /** Where the companion may run for this window. */
  canRun(): CanRun {
    return canRunCompanion({
      trusted: vscode.workspace.isTrusted,
      remoteName: vscode.env.remoteName,
      folderSchemes: (vscode.workspace.workspaceFolders ?? []).map((f) => f.uri.scheme),
      node: process.versions.node,
    });
  }

  /** The agent's name (P43): taken at the companion's next start. */
  name(): string {
    const name = vscode.workspace.getConfiguration('docx4j.agent').get<string>('name', 'Agent').trim();
    return name === '' ? 'Agent' : name;
  }

  /** The companion when it runs. */
  get running(): Running | undefined {
    return this.runner.running;
  }

  /**
   * The workspace's ports, picked (and kept) at the first need: a page's policy names the relay's port when the page is first
   * written, whether or not the companion is running (P37). The running companion's, if it runs.
   */
  async ports(): Promise<CompanionPorts> {
    if (this.runner.running) return this.runner.running.launch.ports;
    const kept = this.context.workspaceState.get<unknown>(PORTS);
    if (isPorts(kept)) return kept;
    // Once: two pages written together in a new workspace must name the same relay.
    this.picking ??= choosePorts(undefined, isFree).then(async ({ ports }) => { await this.context.workspaceState.update(PORTS, ports); return ports; }).finally(() => { this.picking = undefined; });
    return this.picking;
  }

  private picking: Promise<CompanionPorts> | undefined;

  /** The workspace's relay address, for a page's policy. */
  async relay(): Promise<string> {
    return relayAddressOf(await this.ports());
  }

  /** The relay's address as kept, without picking: whether a backup's room is on it. */
  keptRelay(): string | undefined {
    const running = this.runner.running;
    if (running) return running.relay;
    const kept = this.context.workspaceState.get<unknown>(PORTS);
    return isPorts(kept) ? relayAddressOf(kept) : undefined;
  }

  /** The companion started if need be; rejects with words to show. */
  async start(): Promise<Running> {
    const can = this.canRun();
    if (!can.ok) throw new Error(agentWords.cannotRun(can.why));
    if (this.runner.gaveUp) throw new Error(agentWords.gaveUp());
    try {
      return await this.runner.start();
    } catch (e) {
      if (this.runner.gaveUp) throw new Error(agentWords.gaveUp());
      throw new Error(agentWords.failed((e as Error).message));
    }
  }

  /** A page origin allowed at the companion's relay: kept for the next start, and posted to a running companion (P38). */
  async allowOrigin(origin: string): Promise<void> {
    if (!isWebviewOrigin(origin)) return;
    const kept = this.context.globalState.get<string[]>(ORIGINS, []);
    const next = withOrigin(kept, origin);
    if (next.length !== kept.length || next.some((o, i) => o !== kept[i])) await this.context.globalState.update(ORIGINS, next);
    if (this.runner.running && !this.runner.running.launch.origins.includes(origin)) this.runner.post({ type: 'allowOrigin', origin });
  }

  /** The agent invited into a room on the companion's relay; the person has admitted it in this workspace (P40). */
  async invite(room: string, token: string): Promise<void> {
    await this.start();
    await this.admit();
    await this.runner.invite(room, token);
  }

  /** Stopped with the window: `close` posted, nothing killed (section 4.5). */
  close(): void {
    this.runner.close();
  }

  // ---- The process.

  private roots(): string[] {
    return (vscode.workspace.workspaceFolders ?? []).filter((f) => f.uri.scheme === 'file').map((f) => f.uri.fsPath);
  }

  private async token(): Promise<string> {
    let token = await this.context.secrets.get(TOKEN);
    if (!token) {
      token = randomBytes(32).toString('base64url');
      await this.context.secrets.store(TOKEN, token);
    }
    return token;
  }

  private async launch(): Promise<Launch> {
    const kept = this.context.workspaceState.get<unknown>(PORTS);
    const { ports, fresh } = await choosePorts(isPorts(kept) ? kept : undefined, isFree);
    if (fresh) {
      await this.context.workspaceState.update(PORTS, ports);
      if (isPorts(kept)) {
        // The pages written before name the old relay port: their Invite and Share say to reopen the document.
        this.output.appendLine(`port ${kept.port} or ${kept.relay} is taken: the companion moves to ${ports.port} and ${ports.relay}`);
        this.changed.fire();
      }
    }
    const root = this.context.extensionUri.fsPath;
    return {
      cli: join(root, 'companion', 'cli.mjs'),
      ports,
      data: roomsFolder(this.context.storageUri?.fsPath, this.context.globalStorageUri.fsPath, ports, join),
      app: join(root, 'app'),
      roots: this.roots(),
      name: this.name(),
      origins: this.context.globalState.get<string[]>(ORIGINS, []),
      token: await this.token(),
    };
  }

  private spawnChild(launch: Launch): ChildLike {
    this.output.appendLine(`starting on ${launch.ports.port}, its relay on ${launch.ports.relay}`);
    const child = spawn(process.execPath, companionArgs(launch), {
      env: companionEnv(process.env, launch.token),
      stdio: ['ignore', 'ignore', 'pipe', 'ipc'],
      windowsHide: true,
    });
    child.stderr?.setEncoding('utf8');
    child.stderr?.on('data', (text: string) => this.output.append(text));
    return child as unknown as ChildLike;
  }

  // ---- The MCP server (P39, P40).

  private admitted(): boolean {
    return this.context.workspaceState.get<boolean>(ADMITTED, false);
  }

  /** The person admitted the agent in this workspace (an Invite, the address command): VS Code is offered its server from now on. */
  async admit(): Promise<void> {
    if (this.admitted()) return;
    await this.context.workspaceState.update(ADMITTED, true);
    this.changed.fire();
  }

  /** The MCP provider, where VS Code has the API (a fork may lack it, P30). */
  registerMcp(): void {
    const lm = (vscode as { lm?: { registerMcpServerDefinitionProvider?: unknown } }).lm;
    const Definition = (vscode as { McpHttpServerDefinition?: unknown }).McpHttpServerDefinition;
    if (typeof lm?.registerMcpServerDefinitionProvider !== 'function' || typeof Definition !== 'function') return;
    const version = String((this.context.extension.packageJSON as { version?: unknown }).version ?? '0');
    const definition = async (): Promise<vscode.McpHttpServerDefinition> => {
      const ports = await this.ports();
      return new vscode.McpHttpServerDefinition(SERVER_LABEL, vscode.Uri.parse(mcpAddressOf(ports)), { Authorization: `Bearer ${await this.token()}` }, version);
    };
    this.disposables.push(vscode.lm.registerMcpServerDefinitionProvider(MCP_PROVIDER, {
      onDidChangeMcpServerDefinitions: this.changed.event,
      provideMcpServerDefinitions: async () => (offered(this.admitted(), this.canRun()) ? [await definition()] : []),
      resolveMcpServerDefinition: async (server) => {
        if (!(server instanceof vscode.McpHttpServerDefinition)) return server;
        let running = await this.start();
        const token = await this.token();
        // A token changed since the companion started (SecretStorage is the person's, across windows): restarted, or VS Code
        // meets 403, which 1.147 takes for a sign-in it offers to set up by hand (section 9.17).
        if (running.launch.token !== token) running = await this.runner.restart();
        server.uri = vscode.Uri.parse(`${running.url}/mcp`);
        server.headers = { Authorization: `Bearer ${token}` };
        return server;
      },
    }));
  }

  /** "docx4j editor: Copy the Agent's MCP Address": started, admitted, the address copied, Claude Code's line offered. */
  async copyAddress(): Promise<void> {
    const can = this.canRun();
    if (!can.ok) { void vscode.window.showWarningMessage(agentWords.cannotRun(can.why)); return; }
    let running: Running;
    try { running = await this.start(); } catch (e) { void vscode.window.showWarningMessage((e as Error).message); return; }
    await this.admit();
    const address = `${running.url}/mcp`;
    await vscode.env.clipboard.writeText(address);
    const claude = vscode.l10n.t('Copy the Claude Code Command');
    const chosen = await vscode.window.showInformationMessage(vscode.l10n.t('The agent\'s MCP address is copied: {0}. A client sends the token VS Code keeps for you; Claude Code\'s command carries it.', address), claude);
    if (chosen === claude) {
      await vscode.env.clipboard.writeText(claudeCodeLine(address, await this.token()));
      void vscode.window.showInformationMessage(vscode.l10n.t('The Claude Code command is copied. It puts the token in ~/.claude.json and in your shell\'s history.'));
    }
  }
}

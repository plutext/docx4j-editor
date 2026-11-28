#!/usr/bin/env node
// SPDX-License-Identifier: GPL-3.0-or-later
//
// `docx4j-editor-mcp --root <folder> [--root <folder>] [--name Claude] [--port 4174] [--relay-port 4175]
// [--data <folder>] [--app <folder>] [--allow-resolve] [--max-inline-chars <n>] [--open <file>] [--http]
// [--allow-origin <origin>]... [--mcp-token-env <variable>] [--ipc]`: the
// companion (ED-005 section 8.1). It serves the editor, the probe and MCP under `/mcp` on loopback,
// and the relay on a loopback port of its own (a worker thread, ED-005 section 9.3 item 1 (i)), and MCP over stdio for the client that started it (Claude Code: `claude mcp add
// docx4j-editor -- npx @docx4j/editor-mcp --root ~/Documents`), unless `--http` says it runs on its
// own. `--open` shares a file into a room and prints its link on stderr. stdout belongs to MCP.
//
// Started by the VS Code extension (ED-014 section 4.5): `--allow-origin` adds its webview's origin to the local-page rule,
// `--mcp-token-env` names the environment variable holding MCP's token (never an argument, which another user can read),
// and `--ipc` speaks over the channel the spawn opened (`channel.mts`), exiting when it closes.
import { realpathSync } from 'node:fs';
import { fileURLToPath } from 'node:url';
import { format } from 'node:util';
import { StdioServerTransport } from '@modelcontextprotocol/sdk/server/stdio.js';
import { createCompanion, type CompanionOptions } from './index.mjs';
import { invitationRefusal, type Companion } from './server.mjs';
import type { FromCompanion, ToCompanion } from './channel.mjs';

export interface CliOptions extends CompanionOptions {
  open: string[];
  /** MCP over HTTP only, no stdio: the companion started on its own. */
  http: boolean;
  /** The environment variable holding MCP's token (ED-014 P39). */
  mcpTokenEnv?: string;
  /** Speak over the IPC channel the spawn opened (`channel.mts`), and exit when it closes. */
  ipc: boolean;
}

export function parseArguments(argv: readonly string[]): CliOptions {
  const out: CliOptions = { roots: [], open: [], http: false, ipc: false };
  const roots: string[] = [];
  const origins: string[] = [];
  for (let i = 0; i < argv.length; i++) {
    const arg = argv[i]!;
    const value = (): string => {
      const v = argv[++i];
      if (v === undefined) throw new Error(`${arg} needs a value`);
      return v;
    };
    switch (arg) {
      case '--root': roots.push(value()); break;
      case '--name': out.name = value(); break;
      case '--port': out.port = Number(value()); break;
      case '--host': out.host = value(); break;
      case '--relay-port': out.relayPort = Number(value()); break;
      case '--data': out.data = value(); break;
      case '--app': out.app = value(); break;
      case '--allow-resolve': out.allowResolve = true; break;
      case '--open': out.open.push(value()); break;
      case '--http': out.http = true; break;
      case '--allow-origin': origins.push(value()); break;
      case '--mcp-token-env': out.mcpTokenEnv = value(); break;
      case '--ipc': out.ipc = true; break;
      case '--max-inline-chars': out.maxInlineChars = Number(value()); break;
      default: throw new Error(`unknown argument ${arg}`);
    }
  }
  if (out.port !== undefined && !Number.isInteger(out.port)) throw new Error('--port needs a number');
  if (out.relayPort !== undefined && !Number.isInteger(out.relayPort)) throw new Error('--relay-port needs a number');
  if (out.maxInlineChars !== undefined && !(Number.isInteger(out.maxInlineChars) && out.maxInlineChars > 0)) throw new Error('--max-inline-chars needs a positive number');
  if (out.mcpTokenEnv !== undefined && !/^[A-Za-z_][A-Za-z0-9_]*$/.test(out.mcpTokenEnv)) throw new Error('--mcp-token-env needs an environment variable\'s name');
  return { ...out, roots, ...(origins.length > 0 ? { allowOrigins: origins } : {}) };
}

/** MCP's token from the variable `--mcp-token-env` names, taken out of the environment, which a worker would copy. */
export function takeToken(env: NodeJS.ProcessEnv, name: string | undefined): string | undefined {
  if (name === undefined) return undefined;
  const token = env[name];
  if (token === undefined || token.trim() === '') throw new Error(`--mcp-token-env names ${name}, which is not set`);
  delete env[name];
  return token.trim();
}

/**
 * The channel to the process that started the companion (`channel.mts`, ED-014 section 4.5): its messages handled, and its
 * closing (that process gone) or a `close` ending the companion. `stop` writes the rooms and exits.
 */
export function serveChannel(companion: Companion, proc: NodeJS.Process, stop: () => void): (message: FromCompanion) => void {
  const send = (message: FromCompanion): void => { if (proc.connected) proc.send?.(message); };
  proc.on('message', (raw: unknown) => {
    const m = raw as ToCompanion | null;
    if (!m || typeof m !== 'object') return;
    switch (m.type) {
      case 'allowOrigin': if (typeof m.origin === 'string' && m.origin !== '') companion.allowOrigin(m.origin); break;
      case 'roots': if (Array.isArray(m.roots)) companion.setRoots(m.roots.filter((r): r is string => typeof r === 'string')); break;
      case 'invite':
        if (typeof m.id !== 'number' || typeof m.room !== 'string' || typeof m.token !== 'string') return;
        companion.join(m.room, m.token).then(
          () => send({ type: 'invited', id: m.id }),
          (e: unknown) => send({ type: 'invited', id: m.id, ...invitationRefusal(e) }),
        );
        break;
      case 'close': stop(); break;
    }
  });
  // The starter gone: a child with a channel does not exit by itself while its servers hold the event loop (ED-014 section 9.16).
  proc.once('disconnect', stop);
  return send;
}

/**
 * Whether this module is the program node was started with, through any symbolic links: npm installs a bin as a link
 * (`node_modules/.bin/docx4j-editor-mcp` to this file), and npx starts it there, so `process.argv[1]` names the link, not this
 * file. A test that imports the module is not the program.
 */
function startedAsProgram(): boolean {
  const program = process.argv[1];
  if (!program) return false;
  try {
    return realpathSync(program) === realpathSync(fileURLToPath(import.meta.url));
  } catch {
    return false;
  }
}

// Run only as the bin, not when imported by a test.
if (startedAsProgram()) {
  // stdout is the protocol's (vocabulary section 6): whatever a dependency prints with console.log goes to stderr.
  const toStderr = (...args: unknown[]): void => { process.stderr.write(`${format(...args)}\n`); };
  console.log = toStderr;
  console.info = toStderr;
  console.debug = toStderr;
  const options = parseArguments(process.argv.slice(2));
  if (options.ipc && typeof process.send !== 'function') throw new Error('--ipc needs the IPC channel of the process that started the companion');
  const mcpToken = takeToken(process.env, options.mcpTokenEnv);
  let send: (message: FromCompanion) => void = () => undefined;
  const companion = createCompanion({ ...options, ...(mcpToken !== undefined ? { mcpToken } : {}), onRelayState: (state) => send({ type: 'relay', state }) });
  let stopping: Promise<void> | undefined;
  const stop = (): void => {
    stopping ??= companion.close().catch((e: unknown) => { process.stderr.write(`docx4j-editor-mcp: ${String(e)}\n`); }).then(() => process.exit(0));
  };
  if (options.ipc) send = serveChannel(companion, process, stop);
  const url = await companion.listen();
  send({ type: 'listening', url, relay: companion.relay });
  for (const path of options.open) await companion.open(path);
  if (!options.http) await companion.mcp().connect(new StdioServerTransport());
  process.on('SIGINT', stop);
  process.on('SIGTERM', stop);
}

#!/usr/bin/env node
// SPDX-License-Identifier: GPL-3.0-or-later
//
// `docx4j-editor-mcp --root <folder> [--root <folder>] [--name Claude] [--port 4174] [--relay-port 4175]
// [--data <folder>] [--app <folder>] [--allow-resolve] [--max-inline-chars <n>] [--open <file>] [--http]`: the
// companion (ED-005 section 8.1). It serves the editor, the probe and MCP under `/mcp` on loopback,
// and the relay on a loopback port of its own (a worker thread, ED-005 section 9.3 item 1 (i)), and MCP over stdio for the client that started it (Claude Code: `claude mcp add
// docx4j-editor -- npx @docx4j/editor-mcp --root ~/Documents`), unless `--http` says it runs on its
// own. `--open` shares a file into a room and prints its link on stderr. stdout belongs to MCP.
import { format } from 'node:util';
import { StdioServerTransport } from '@modelcontextprotocol/sdk/server/stdio.js';
import { createCompanion, type CompanionOptions } from './index.mjs';

export interface CliOptions extends CompanionOptions {
  open: string[];
  /** MCP over HTTP only, no stdio: the companion started on its own. */
  http: boolean;
}

export function parseArguments(argv: readonly string[]): CliOptions {
  const out: CliOptions = { roots: [], open: [], http: false };
  const roots: string[] = [];
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
      case '--max-inline-chars': out.maxInlineChars = Number(value()); break;
      default: throw new Error(`unknown argument ${arg}`);
    }
  }
  if (out.port !== undefined && !Number.isInteger(out.port)) throw new Error('--port needs a number');
  if (out.relayPort !== undefined && !Number.isInteger(out.relayPort)) throw new Error('--relay-port needs a number');
  if (out.maxInlineChars !== undefined && !(Number.isInteger(out.maxInlineChars) && out.maxInlineChars > 0)) throw new Error('--max-inline-chars needs a positive number');
  return { ...out, roots };
}

// Run only as the bin, not when imported by a test.
if (process.argv[1] && import.meta.url === new URL(`file://${process.argv[1]}`).href) {
  // stdout is the protocol's (vocabulary section 6): whatever a dependency prints with console.log goes to stderr.
  const toStderr = (...args: unknown[]): void => { process.stderr.write(`${format(...args)}\n`); };
  console.log = toStderr;
  console.info = toStderr;
  console.debug = toStderr;
  const options = parseArguments(process.argv.slice(2));
  const companion = createCompanion(options);
  await companion.listen();
  for (const path of options.open) await companion.open(path);
  if (!options.http) await companion.mcp().connect(new StdioServerTransport());
  const stop = (): void => { void companion.close().then(() => process.exit(0)); };
  process.on('SIGINT', stop);
  process.on('SIGTERM', stop);
}

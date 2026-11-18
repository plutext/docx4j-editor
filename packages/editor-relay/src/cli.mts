#!/usr/bin/env node
// SPDX-License-Identifier: GPL-3.0-or-later
//
// docx4j-editor-relay: the relay on its own (ED-004 section 4.2).
//
//   docx4j-editor-relay [--port 4175] [--host 127.0.0.1] [--data ./rooms]
//                       [--max-room 209715200] [--idle-days 30] [--origin <url>]...
import { createRelay, type RelayOptions } from './index.mjs';

const USAGE = `usage: docx4j-editor-relay [--port n] [--host address] [--data folder] [--max-room bytes] [--idle-days n] [--origin url]...

  --port       the port to listen on (default 4175)
  --host       the address to bind (default 127.0.0.1: this machine only)
  --data       the folder of room files (default ./rooms)
  --max-room   a room's largest size in bytes (default 209715200)
  --idle-days  days without a change before a room is deleted (default 30)
  --origin     a page origin allowed to connect, repeatable; '*' for any
               (default: pages on localhost, 127.0.0.1 or [::1])`;

export function parseArguments(argv: readonly string[]): RelayOptions {
  const options: RelayOptions = {};
  const origins: string[] = [];
  for (let i = 0; i < argv.length; i++) {
    const flag = argv[i]!;
    const value = argv[i + 1];
    const need = (): string => {
      if (value === undefined || value.startsWith('--')) throw new Error(`${flag} needs a value`);
      i++;
      return value;
    };
    const number = (): number => {
      const n = Number(need());
      if (!Number.isFinite(n) || n < 0) throw new Error(`${flag} needs a number`);
      return n;
    };
    switch (flag) {
      case '--port': options.port = number(); break;
      case '--host': options.host = need(); break;
      case '--data': options.data = need(); break;
      case '--max-room': options.maxRoom = number(); break;
      case '--idle-days': options.idleDays = number(); break;
      case '--origin': origins.push(need()); break;
      default: throw new Error(`unknown option ${flag}`);
    }
  }
  if (origins.length > 0) options.origins = origins;
  return options;
}

async function main(): Promise<void> {
  const argv = process.argv.slice(2);
  if (argv.includes('--help') || argv.includes('-h')) { console.log(USAGE); return; }
  let options: RelayOptions;
  try {
    options = parseArguments(argv);
  } catch (e) {
    console.error(`${(e as Error).message}\n\n${USAGE}`);
    process.exitCode = 2;
    return;
  }
  const relay = createRelay(options);
  const url = await relay.listen();
  console.log(`docx4j-editor-relay listening on ${url} (rooms in ${options.data ?? './rooms'})`);
  const stop = (): void => { void relay.destroy().then(() => process.exit(0)); };
  process.once('SIGINT', stop);
  process.once('SIGTERM', stop);
}

// Run when executed, not when imported (the tests import `parseArguments`).
if (import.meta.url === `file://${process.argv[1]}` || process.argv[1]?.endsWith('docx4j-editor-relay')) void main();

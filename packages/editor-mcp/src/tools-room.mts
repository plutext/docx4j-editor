// SPDX-License-Identifier: GPL-3.0-or-later
//
// The room's tools (docs/developer/agent.md section 3.1): open a file into a room, list the rooms, leave one,
// save one, and report the tracking (which the agent never turns off).
import { writeFile } from 'node:fs/promises';
import type { McpServer } from '@modelcontextprotocol/sdk/server/mcp.js';
import { z } from 'zod';
import { ToolError } from './errors.mjs';
import { writableUnderRoots } from './paths.mjs';
import { answer, documentArg, inRoom, ok, type ToolHost } from './kit.mjs';

export function registerRoomTools(server: McpServer, host: ToolHost): void {
  server.registerTool('open', {
    title: 'Open a document',
    description: 'Opens a .docx under one of the companion\'s --root folders in a co-editing room, the agent joined, and returns the room id (the document handle every other tool takes) and the link a person opens to join in the editor. A file already open answers its room. A file whose document the person stopped you in (its own, or a copy you saved) answers agent.stopped until they invite you back.',
    inputSchema: { path: z.string().min(1).describe('the file, absolute or relative to the first root') },
  }, (args) => answer(async () => {
    const room = await host.open(args.path);
    // A file read through Word's repair (docs/developer/agent.md section 3.1): the load's own warnings, each naming what was repaired.
    const pkg = room.session.pkg;
    const repaired = pkg.repaired === 'none' ? [] : pkg.warnings.filter((w) => w.code.startsWith('package.')).map((w) => w.message);
    const said = repaired.length > 0 ? ` It was read as Word repairs it, and a save writes the repaired file: ${repaired.join('; ')}.` : '';
    return ok({ document: room.id, link: room.link, path: room.path ?? null, ...(repaired.length > 0 ? { repaired } : {}) }, `Opened ${room.path ?? args.path} as ${room.id}.${said} Give the person this link to join: ${room.link}`);
  }));

  server.registerTool('list_documents', {
    title: 'List the open documents',
    description: 'The rooms the agent is in: each document handle, its file, its link and who is in it.',
    inputSchema: {},
  }, () => answer(async () => {
    const documents = [...host.rooms.values()].map((room) => ({ document: room.id, path: room.path ?? null, link: room.link, people: room.people(), ...(room.stopped ? { stopped: true } : {}) }));
    return ok({ documents });
  }));

  server.registerTool('close', {
    title: 'Leave a document\'s room',
    description: 'Leaves the room; the person\'s editor stays open, and nothing is saved.',
    inputSchema: { document: documentArg },
  }, (args) => answer(async () => {
    const room = host.rooms.get(args.document);
    if (!room) throw new ToolError('document.not_open', `no document is open as ${args.document}`, 'call list_documents', { document: args.document });
    await room.close();
    return ok({ document: args.document }, `Left ${args.document}`);
  }));

  server.registerTool('save', {
    title: 'Save a document',
    description: 'Writes the document as the room has it, every tracked change pending as it is, to its own file, or to output_path under a root (overwrite: true to replace an existing file other than its own).',
    inputSchema: {
      document: documentArg,
      output_path: z.string().min(1).optional(),
      overwrite: z.boolean().optional(),
    },
  }, (args) => answer(() => inRoom(host, args.document, async (room) => {
    let target = room.path;
    if (target === undefined && args.output_path === undefined) {
      throw new ToolError('arguments.invalid', 'this document came from a person\'s browser and has no file of its own', 'give output_path, under a root', { argument: 'output_path' });
    }
    if (args.output_path !== undefined) {
      const w = await writableUnderRoots(host.roots, args.output_path);
      if (w.exists && w.path !== room.path && args.overwrite !== true) {
        throw new ToolError('path.exists', `${args.output_path} exists`, 'pass overwrite: true to replace it, or choose another name', { path: w.path });
      }
      target = w.path;
    }
    const saved = await room.session.save();
    await writeFile(target!, saved.bytes);
    room.saved.add(target!);
    return ok({ path: target, bytes: saved.bytes.length }, `Saved ${target} (${saved.bytes.length} bytes)`);
  })));

  server.registerTool('track_changes', {
    title: 'Tracking',
    description: 'Reports the agent\'s tracking (always on: every edit it makes is a tracked change under its name) and the document\'s own setting, which the agent never changes. mode "Off" is refused.',
    inputSchema: { document: documentArg, mode: z.enum(['TrackAll', 'Off']).optional() },
  }, (args) => answer(() => inRoom(host, args.document, async (room) => {
    if (args.mode === 'Off') throw new ToolError('tracking.not_allowed', 'the agent\'s edits are always tracked', 'every edit you make waits for the person to accept or reject it; that is the point of the room', {});
    return ok({ mode: 'TrackAll', author: host.name, document_mode: room.session.trackingMode });
  })));
}

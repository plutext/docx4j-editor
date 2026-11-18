// SPDX-License-Identifier: GPL-3.0-or-later
//
// Accept and reject (docs/developer/agent.md section 3.5), only with `--allow-resolve`: an agent resolving its
// own changes undoes the review the tracking exists for (ED-005 proposal 12). Every resolution goes
// through the four commands, as the review panel's does, so a room's repair sees it (section 12.15)
// and the revisions shown in Word only go to the engine in their settle half (section 12.26).
import type { McpServer } from '@modelcontextprotocol/sdk/server/mcp.js';
import { z } from 'zod';
import { runCommandAsync } from '@docx4j/editor-model';
import { ToolError } from './errors.mjs';
import { addressSchema, resolveBlock } from './address.mjs';
import { answer, documentArg, inRoom, ok, type ToolHost, type ToolRoom } from './kit.mjs';
import { listChanges, PRESERVED_KEY, type ListedChange } from './tools-read.mjs';

/** One change by its key, through `trackedChange.accept` or `.reject`; false when it is gone. */
async function resolveKey(room: ToolRoom, action: 'accept' | 'reject', key: string): Promise<boolean> {
  const args = key.startsWith(PRESERVED_KEY) ? { preserved: JSON.parse(key.slice(PRESERVED_KEY.length)) as unknown } : { key };
  const result = await runCommandAsync(room.session, `trackedChange.${action}`, args);
  return result.ok;
}

export function registerResolveTools(server: McpServer, host: ToolHost): void {
  for (const action of ['accept', 'reject'] as const) {
    server.registerTool(`${action}_tracked_changes`, {
      title: action === 'accept' ? 'Accept tracked changes' : 'Reject tracked changes',
      description: `${action === 'accept' ? 'Accepts' : 'Rejects'} the tracked change a key names (get_tracked_changes gives them), those at an address, or all (an author's with author), as the review panel does. Only when the companion was started with --allow-resolve; otherwise refused with resolve.not_allowed.`,
      inputSchema: { document: documentArg, key: z.string().min(1).optional(), address: addressSchema.optional(), author: z.string().optional() },
    }, (args) => answer(() => {
      if (!host.allowResolve) throw new ToolError('resolve.not_allowed', 'the agent may not accept or reject tracked changes here', 'the person reviews your changes; the companion must be started with --allow-resolve for this', {});
      if (args.key !== undefined && args.address !== undefined) throw new ToolError('arguments.invalid', 'give a key or an address, not both', 'a key names one change; an address takes those in a block', { argument: 'key' });
      return inRoom(host, args.document, async (room) => {
        const verb = action === 'accept' ? 'Accepted' : 'Rejected';
        if (args.key !== undefined) {
          if (!(await resolveKey(room, action, args.key))) throw new ToolError('target.stale', 'the change the key names is no longer there (resolved, or its text deleted)', 'call get_tracked_changes for the changes as they are now', { key: args.key });
          return ok({ summary: `${verb} one change`, resolved: 1 });
        }
        // The changes the listing gives, in the scope and by the author asked for.
        const listed: ListedChange[] = await room.session.readTree((b) => {
          const scope = args.address !== undefined ? b.body.addressOf(resolveBlock(b.body, args.address).block) : undefined;
          return listChanges(room, b).filter((c) => (scope === undefined || (c.address !== null && (c.address === scope || c.address.startsWith(`${scope}/`)))) && (args.author === undefined || c.author === args.author));
        });
        if (args.address === undefined) {
          // All, or all of an author's: the command resolves what the listing has, the engine what only "all" reaches.
          if (listed.length === 0) return ok({ summary: `No tracked changes to ${action}`, resolved: 0 });
          const result = await runCommandAsync(room.session, action === 'accept' ? 'trackedChanges.acceptAll' : 'trackedChanges.rejectAll', args.author !== undefined ? { author: args.author } : undefined);
          return ok({ summary: `${verb} ${listed.length === 1 ? 'one change' : `${listed.length} changes`}`, resolved: result.ok ? listed.length : 0 });
        }
        let resolved = 0;
        for (const c of listed) if (c.key !== undefined && (await resolveKey(room, action, c.key))) resolved++;
        const left = listed.filter((c) => c.key === undefined).length;
        return ok({ summary: `${verb} ${resolved === 1 ? 'one change' : `${resolved} changes`} at ${String(args.address)}`, resolved, ...(left > 0 ? { warnings: [`${left} changes there are resolved only by all, since Word's API does not list them (a cell merge, a list's numbering)`] } : {}) });
      });
    }));
  }
}

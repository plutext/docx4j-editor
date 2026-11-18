# @docx4j/editor-mcp

The companion for [`@docx4j/editor`](../editor): one local process that serves the editor, its
co-editing relay and MCP tools, so that an agent edits a document in a room beside a person, every
edit a tracked change waiting for accept or reject. The tools and their contract are
[`docs/developer/agent.md`](../../docs/developer/agent.md).

## Setting it up

An MCP client starts it over stdio.

**The package is not on npm yet** (publishing is a release question). Until it is, run it from a
checkout of this repository: `npm install` and `npm run build` at the repository's root, then give
the client `node <checkout>/packages/editor-mcp/dist/cli.mjs` where the lines below say
`npx @docx4j/editor-mcp`. From the root, `npm run companion -- --root <folder> --open <file>` builds
and starts it on its own.

**Claude Code:**

```
claude mcp add docx4j-editor -- npx @docx4j/editor-mcp --root ~/Documents
```

From a checkout:

```
claude mcp add docx4j-editor -- node ~/git/docx4j-ts-editor/packages/editor-mcp/dist/cli.mjs --root ~/Documents
```

`claude mcp list` shows it. To try it without touching your own configuration, give Claude Code a
file of its own: `claude --strict-mcp-config --mcp-config companion.json`, the file holding the
`mcpServers` entry below.

**Claude Desktop:** add the server to `claude_desktop_config.json` (Settings, Developer, Edit Config)
and restart Claude Desktop. This entry follows Anthropic's documentation and is untested here, since
Claude Desktop has no Linux build:

```json
{
  "mcpServers": {
    "docx4j-editor": {
      "command": "npx",
      "args": ["-y", "@docx4j/editor-mcp", "--root", "C:\\Users\\you\\Documents"]
    }
  }
}
```

From a checkout, `"command": "node"` and `"args": ["C:\\path\\to\\docx4j-ts-editor\\packages\\editor-mcp\\dist\\cli.mjs", "--root", "..."]`.

**Loopback only.** The companion binds `127.0.0.1` and refuses a request whose `Host` is anything
else, so the editor, the relay and MCP over HTTP are reached from a browser on the same machine only.
A browser in a virtual machine, or on another computer, does not reach it: open the room link on the
machine that runs the companion. It is not a server to expose.

**A rebuilt editor.** The companion serves its own copy of the editor, copied into its `dist/app` when
it is built. After rebuilding the editor, rebuild the companion (`npm run build -w @docx4j/editor-mcp`)
and restart it: a running companion keeps serving the copy it started with. `--app` serves another
folder instead.

## Using it

The agent's `open` answers a room link to give the person; opening it in a browser on the same
machine joins the room in the editor. On its own, `--http` serves MCP over streamable HTTP at
`http://127.0.0.1:4174/mcp` instead of stdio, and `--open` shares a file at start and prints its
link on stderr:

```
npx @docx4j/editor-mcp --root ~/Documents --http --open lease.docx
```

stdout stays silent: it belongs to MCP.

| Option | Default | |
|---|---|---|
| `--root` | none | a folder the agent may open files under, repeatable; symlinks are resolved, and a path outside every root is refused |
| `--name` | `Claude` | the agent's name, on its tracked changes and comments and in the People list |
| `--port` | 4174 | the port for the editor (`/`), the probe (`/companion/`) and MCP over HTTP (`/mcp`) |
| `--relay-port` | the port and one (4175) | the relay's own port: it runs in a worker thread, so people's typing to each other flows while an agent's call runs; the room links carry it. A default that is taken falls back to a free port |
| `--host` | `127.0.0.1` | the address to bind; requests with any other `Host` are refused (DNS rebinding) |
| `--data` | `./.docx4j-editor-rooms` | the relay's room files |
| `--app` | the built editor | the application's folder, when not the package's own `dist/app` |
| `--allow-resolve` | off | let the agent accept and reject tracked changes (ED-005 proposal 12: off, so the agent cannot undo the review) |
| `--open` | none | share a file into a room at start and print its link, repeatable |
| `--http` | off | MCP over streamable HTTP under `/mcp` only, not stdio: the companion started on its own |
| `--max-inline-chars` | 200000 | the cap on a result's characters; past it a result is truncated, or written to the call's `output_path` |

**The agent is always tracked**, whatever the document's own setting, and never changes that
setting (ED-005 section 8.5). The agent joins each room as a page does, over loopback to the
relay, and its presence is its connection's awareness state.

**The browser first.** A page the companion serves (open `http://127.0.0.1:4174`) finds it at its
own `/companion`, and offers **Invite Claude**: the document is shared into the companion's relay and
the companion told the room (`POST /companion/rooms` with the room and its token), and the agent joins.
In the People list the agent is badged and has a **Stop**, which any person in the room may press: it
is written into the room, and the companion leaves and refuses the agent's further calls there. What
the agent did is logged in the room, and shown in the review panel.

The tools, by the portfolio vocabulary's names:

| | |
|---|---|
| The room | `open`, `list_documents`, `close`, `save`, `track_changes` |
| Reading | `outline`, `find`, `read` (Markdown with addresses), `extract_text` (plain text), each through the text model with the text extraction specification's options (`view`, `include` and the rest; `docs/developer/agent.md` section 2.13), `get_comments`, `get_tracked_changes` |
| Editing | `edit`, `insert`, `delete`, `replace_text`, `set_style`, `set_format`, `insert_ooxml`, `comment` |
| Content controls | `insert_content_control`, `remove_content_control`, `set_content_control`, `bind` |
| Resolving (`--allow-resolve`) | `accept_tracked_changes`, `reject_tracked_changes` |

No tool runs a script or a macro, and none edits a package part as XML.

GPL-3.0-or-later, as the rest of the repository; a commercial licence is available.

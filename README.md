# docx4j editor

A Word docx document editor in the browser, including:

- **Review:** tracked changes, written as Word writes them, with a review panel to accept or reject.
- **Comments:** made, replied to, resolved and deleted in a margin.
- **Developer tools:** the document's parts and XML (editable), a console for the content API, and macros.
- **An AI agent as a peer,** every edit it makes a tracked change.
- **Content controls:** a Template perspective for OpenDoPE templates.
- **Co-editing:** several people in one document through a relay.

The application is a static site: it runs entirely in the browser and needs no server of its own.

It opens a `.docx` or `.md` and saves `.docx` back. A document opened and saved without edits comes back
identical part by part, and an edit changes only what it touched.

Built on [`@docx4j/core-ts`](https://github.com/plutext/docx4j-core-ts).

Ctrl+S saves (in place when the
file was opened through the browser's file picker and that option is on, else as a download), and
Ctrl+P prints. The status bar's zoom (or Ctrl+wheel over the page) enlarges the page alone; the
browser's Ctrl+plus enlarges everything.

## Alternatives

As of November 2026:

- [docx-editor.dev](https://docx-editor.dev/) by EigenPal is open core: its core and React editor are Apache-2.0,
  but comments, tracked changes and real-time collaboration are in `@docx-editor.dev/pro`, under a commercial licence.
- [SuperDoc](https://github.com/superdoc-dev/superdoc) is AGPL-3.0, but its current releases (2.33) depend on
  `@superdoc/docx-engine`, a closed-source package under a proprietary licence. SuperDoc 1 is superseded.

## Requirements

- **Node.js 22.12 or later** (22, 24 and 26 are tested). Vite 7, which builds and serves the application, refuses
  older versions of 22. `.nvmrc` names 22, so with nvm, `nvm install` and then `nvm use` in the repository's folder are
  enough.
- **npm**, which comes with Node.js.

Check with `node --version` before you start.

## Starting the editor

From the repository's root:

```
npm install
npm run dev
```

`npm run dev` builds the two packages the application imports (`editor-model` and `editor-relay`),
then starts Vite's development server. Open the address it prints, `http://localhost:5173/`, and
choose **Open…** in the application bar to open a `.docx`.

**From another machine or a virtual machine.** Vite's server listens on this machine only. To reach
it from elsewhere, for example from a VirtualBox guest, start it with `--host`:

```
npm run dev -w @docx4j/editor -- --host
```

Then open the host's address from the guest, for example `http://10.0.2.2:5173/`. This form starts
Vite alone, so run `npm run dev` once first, and again after changing `packages/editor-model` or
`packages/editor-relay`. A page opened this way edits and saves as usual. The relay the development
server carries accepts pages on this machine only, so sharing a document from it is refused.

**The built application.** To build it and serve the result:

```
npm run build
npm run preview -w @docx4j/editor
```

`npm run preview` serves `packages/editor/dist` at `http://localhost:4173/`. That folder is the site
to deploy: any static web server can serve it.

## Sharing a document

The development and preview servers carry a co-editing relay under `/relay`.

1. Set **Options → Relay address** to `ws://localhost:5173/relay` (`ws://localhost:4173/relay` on the
   preview server).
2. Choose **Share**, and give the link it makes to the other person.

To run the relay on its own, start it with `npm run relay`. It listens on `ws://127.0.0.1:4175`, the
address to give, and keeps its rooms in `rooms/` in the repository.
`docs/user/co-editing.md` covers sharing in full, and `packages/editor-relay/README.md` covers the
relay's options and deployment.

## An agent in the document

The companion serves the editor, a relay and the agent's MCP tools as one local process. It is on
npm as [`@docx4j/editor-mcp`](https://www.npmjs.com/package/@docx4j/editor-mcp), so it needs no clone. To
start it with a file shared:

```
npx @docx4j/editor-mcp --root <folder> --open <file>
```

From a clone of this repository, `npm run companion -- --root <folder> --open <file>` builds it and
starts it the same way.

It listens on `http://127.0.0.1:4174/`. `packages/editor-mcp/README.md` covers connecting Claude
Code or Claude Desktop to it, and `docs/developer/agent.md` describes the tools.

## Developing

| Command | |
|---|---|
| `npm run typecheck` | TypeScript, strict, across every package |
| `npm test` | the unit and integration suites (vitest) in every package. The model's takes under two minutes, most of it its benchmark, which reports figures and never fails |
| `npm run e2e` | the browser suite (Playwright, Chromium) against the built application; once, first: `npm run e2e:install -w @docx4j/editor` |
| `npm run locales:check` | the message files of the twenty languages, against the code and each other |
| `npm run check-licences` | every dependency under a permissive licence (MIT, BSD, Apache-2.0); it also writes each npm package's `THIRD-PARTY-NOTICES.txt`, and the site's build writes the site's |

In one package: `npx vitest run [file] [-t name]` in its folder.

The packages, under `packages/`:

- **`editor-model`** (`@docx4j/editor-model`): the document as an editable ProseMirror document over Yjs.
  It holds the commands, tracked changes and comments, and runs headless in Node.
- **`editor`** (`@docx4j/editor`): the React application and its view.
- **`editor-relay`** (`@docx4j/editor-relay`): the co-editing relay, a Hocuspocus server.
- **`editor-mcp`** (`@docx4j/editor-mcp`): the agent's companion.

## Documentation

For people using the editor, in [`docs/user/`](docs/user):

- [`review.md`](docs/user/review.md): reviewing, tracking changes, comments and the agent;
- [`co-editing.md`](docs/user/co-editing.md): sharing a document;
- [`files.md`](docs/user/files.md): opening, saving, and a file changed on disk;
- [`printing.md`](docs/user/printing.md);
- [`developer-perspective.md`](docs/user/developer-perspective.md): the Developer perspective, and
  [`Help_AskBox.md`](docs/user/Help_AskBox.md), its Ask box.

For contributors, in [`docs/developer/`](docs/developer):

- [`architecture.md`](docs/developer/architecture.md): the three layers and the rules the projection keeps;
- [`agent.md`](docs/developer/agent.md): the agent's tools;
- [`localisation.md`](docs/developer/localisation.md): the interface's languages;
- [`word-divergences.md`](docs/developer/word-divergences.md): where the editor knowingly differs from Word;
- [`upstream-limitations.md`](docs/developer/upstream-limitations.md): third-party limitations and their
  workarounds;
- [`TODOs_OpenDoPE.md`](docs/developer/TODOs_OpenDoPE.md): what OpenDoPE templates do that the editor does
  not yet.

## Licence

`GPL-3.0-or-later` (see `LICENSE`), with a commercial licence available alongside. Contributions come
in under Apache-2.0 (see `CONTRIBUTING.md`).

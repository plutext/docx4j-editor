# Security

## Reporting a vulnerability

Please report a vulnerability privately, not in a public issue: use the **Report a vulnerability** button on the
repository's [Security tab](https://github.com/plutext/docx4j-editor/security/advisories/new), GitHub's private
vulnerability reporting.

Say what an attacker can do, which version or commit you used, and how to reproduce it. A file that shows the
problem is welcome; remove anything private from it first.

We'll acknowledge the report, work on a fix with you, and publish an advisory when the fix is released, crediting
you unless you'd rather not be named.

## Supported versions

The editor is at 0.1, a preview. Fixes go into `main` and the next release; earlier releases get none.

## What counts

The editor opens files from anywhere, so these are vulnerabilities:

- a document (`.docx`, `.md`, Flat OPC `.xml`) or pasted content that runs script in the page, makes the browser
  fetch anything, or reads anything beyond itself, when it is opened, edited or saved;
- a macro in a document that runs without the person choosing to run it, or without being shown to them first;
- the Ask box's API key, kept in the browser's local storage, going anywhere but the provider address the person set;
- the relay (`@docx4j/editor-relay`) letting a connection into a room without the room's token, or letting a
  read-only connection write;
- the companion (`@docx4j/editor-mcp`) answering a request from another machine or another origin, or opening a file
  outside its `--root` folders;
- one of the agent's tools running a script or a macro, or making an edit that is not a tracked change.

## By design

- The console runs the JavaScript you type, in the page: that is what it is for. A macro you choose to run runs
  the same way. A deployment can turn document macros off with `VITE_DOCUMENT_MACROS=off` at build time.
- The Ask box sends your question, with the parts of the document it describes, to the provider you configure,
  using your key.
- A relay reached over plain `ws:` from another machine carries the room's token in clear: put a relay that others
  reach behind TLS (`packages/editor-relay/README.md`).

## Dependencies

A vulnerability in `@docx4j/core-ts`, the document engine, or in another dependency can be reported here too; we
will pass it on.

# Contributing to the docx4j editor

Thank you for your interest in contributing.

## Licensing: contributions come in under Apache-2.0

**Contributions to every package in this repository are accepted under the
[Apache License, Version 2.0](https://www.apache.org/licenses/LICENSE-2.0),
whichever package they touch.** This is deliberate, and it differs from
`docx4j`, where inbound and outbound are the same licence. Please read this
section before your first pull request.

What the project distributes is `GPL-3.0-or-later`, in every package here:

| Package | Distributed under |
|---|---|
| `@docx4j/editor-model` (headless projection and commands) | `GPL-3.0-or-later` |
| `@docx4j/editor-mcp` (the MCP companion) | `GPL-3.0-or-later` |
| `@docx4j/editor` (the assembled application) | `GPL-3.0-or-later` |

A commercial licence is available alongside, for anyone embedding the editor in
a product they do not wish to publish the source of. The engine this builds on,
`@docx4j/core-ts`, is Apache-2.0 and lives in its own repository.

So the inbound licence (Apache-2.0) is more permissive than the outbound one
(GPL). **In one plain sentence: you grant permissive terms,
you receive copyleft, and Plutext may ship your work in a commercially licensed
build sold to customers who do not want the GPL.** That arrangement is
well-precedented in dual-licensed projects, but it is the sort of thing people
are reasonably annoyed to discover late, so it is stated here rather than in a
footnote. Two consequences worth knowing: your contribution stays Apache-2.0
permanently, so anyone may extract those parts under Apache-2.0 terms; and
Apache §3 means you grant a patent licence with it.

Source files carry `SPDX-License-Identifier: GPL-3.0-or-later`, and the inbound
licence above is what your contribution is offered under.

**Dependencies must be permissive (MIT, BSD, Apache-2.0) in every package**,
even though this repository is GPL. A copyleft dependency owned by someone else
would remove the project's ability to offer the commercial licence at all, and
it cannot be undone later without replacing the dependency. If a change adds a
dependency, name its licence in the pull request.

**Editor code may not move down into `@docx4j/core-ts` or
`@docx4j/generated-objects-ts`.** Those are Apache-2.0, in their own
repository, and have to stay usable by Apache-2.0 consumers such as Office
add-ins. Moving code the other way, engine into editor, is fine.

If you are not willing to contribute on these terms, please open an issue to
discuss it before writing code, rather than sending a pull request that cannot
be merged.

## Certifying your contribution (DCO)

This project uses the [Developer Certificate of Origin](https://developercertificate.org/)
(DCO). By adding a `Signed-off-by:` line to each commit, you certify that you
wrote the contribution, or otherwise have the right to submit it under the
Apache License v2, as set out in the DCO text.

Add the sign-off with the `-s` flag when committing:

```bash
git commit -s -m "Keep sdt groups contiguous when a row is deleted (issue 123)"
```

which appends a trailer of the form:

```
Signed-off-by: Your Name <your.email@example.com>
```

Use your real name and a working email address. The sign-off must be added by
a human: automated tools and AI agents must not add `Signed-off-by:` lines
(see below).

Pull requests are checked for a valid sign-off on every commit. If you forgot
it, you can amend with `git commit --amend -s` or
`git rebase --signoff HEAD~<n>`.

## AI-assisted contributions

Contributions written with the help of AI tools (Claude, Copilot, and similar)
are welcome, subject to the following conditions:

1. **You are responsible for the contribution.** You must have personally
   reviewed and understood every part of the change, and be able to explain
   and defend it in code review. Bugs, regressions, licensing problems, and
   provenance issues in AI-assisted code remain your responsibility, exactly
   as if you had typed it yourself. Pull requests whose submitter cannot
   answer questions about the code will be closed.

2. **Tool terms must be compatible.** The terms of use of the AI tool must not
   place restrictions on its output that are inconsistent with the Apache
   License v2 or the Open Source Definition.

3. **No third-party code laundering.** Do not submit AI output that you know
   or suspect reproduces identifiable third-party code. If your tool flags
   output as matching existing licensed code, either comply with that license
   (including any NOTICE/attribution obligations) or do not submit it.

4. **Disclose AI assistance** with a commit trailer. Preferred form:

   ```
   Assisted-by: Claude Code (claude-fable-5)
   ```

   `Generated-by:` and `Co-Authored-By:` trailers are also accepted. Purely
   mechanical assistance (autocomplete of a line or two) does not need
   disclosure; generation of substantive logic does.

5. **Only humans sign off.** The `Signed-off-by:` line certifies the DCO and
   must be added by you, not by the tool. Do not configure an AI agent to add
   sign-offs on your behalf.

## Practical guidelines

- **Discuss first**: for anything beyond a small fix, please open an issue
  describing the problem before investing effort in a pull request.
  `docs/developer/architecture.md` describes the design: the three layers and the rules
  the projection keeps.
- **The projection is where the fidelity lives.** Changes to import or export
  need a round-trip test over the fixtures: an untouched document must export
  byte for byte, and an edited one must differ only in the edit. See
  `docs/developer/architecture.md`.
- **Translations**: adding a language, reviewing a draft and adding a string to the
  interface are described in `docs/developer/localisation.md`.
- **Style**: match the conventions of the surrounding code; use Unix (LF)
  line endings.

The commands that build, check and test the packages are in the README's
"Developing" table.

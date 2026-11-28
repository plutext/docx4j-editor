# The VS Code extension, for developers

The extension (`packages/editor-vscode`, the npm workspace `docx4j-editor`) puts the editor in VS Code as a custom editor
for Word documents. Its README is the listing's page, for people; this page is for working on it. The design is
`docs/change-requests/ED-014-vscode-extension.md`.

## What it is made of

- **`dist/extension.js`:** the extension's code, `src/`, bundled by esbuild into one CommonJS file
  (`scripts/build.mjs`). It takes the application's `host-protocol.ts`, `webview-page.ts` and `ask-send.ts` through
  `src/shared.ts`, and the Anthropic SDK with them. Beside the manifest, the build writes `THIRD-PARTY-NOTICES.txt` for the
  packages it bundled.
- **`app/`:** the application's `dist` without its source maps (`scripts/copy-app.mjs`), which the webviews load. It
  carries its own `LICENSE.txt` and `THIRD-PARTY-NOTICES.txt`.
- **`companion/`:** the agent's companion, `@docx4j/editor-mcp`'s `dist` bundled by esbuild with every package it imports
  (`scripts/bundle-companion.mjs`), and its notices. The extension spawns it under VS Code's own Node. Its output channel is
  "docx4j editor: Agent".
- **The words:** `package.nls.json` and `l10n/bundle.l10n.json`, English, with a draft of each per language
  (`docs/developer/localisation.md`).

The extension has no runtime dependency, so the VSIX holds no `node_modules`.

## Building and testing

From the repository's root, after `npm install`:

```
npm run build                                       # every package; this one last: dist/, app/ and companion/
npm run test -w docx4j-editor                       # the unit tests (vitest), over the modules with no vscode import
xvfb-run -a npm run test:vscode -w docx4j-editor    # the suite in a real VS Code
npm run vsix                                        # the build, then the VSIX
```

**The suite in VS Code** drives a real VS Code over the DevTools Protocol with Playwright, on a scratch profile per spec.
A driver extension (`test/vscode/driver/`) runs VS Code's commands and reads its tabs and clipboard for it.
- **Which VS Code:** it downloads 1.147.0 into `.vscode-test/` the first time. `DOCX4J_VSCODE` names an executable to use
  instead, and `DOCX4J_VSCODE_VERSION` another version to download; `DOCX4J_VSCODE_VERSION=1.105.0` runs the floor.
- **A display:** on Linux without one, run it under `xvfb-run -a`. On a desktop it opens windows on the screen. VS Code
  1.105.0 drops the suite's clicks while the desktop session is locked.
- **`installed.spec.ts`** packages the extension, the driver and a stand-in French language pack, installs them through VS
  Code's command line, and runs a window with no development path, as a person's is.
- **`DOCX4J_SUITE_DEBUG=1`** logs VS Code's arguments and the DevTools connection's attempts.

**By hand:** run VS Code with `--extensionDevelopmentPath=<the repository>/packages/editor-vscode`. On Linux under Wayland
with an NVIDIA driver, add `--ozone-platform=x11`.

## The VSIX

`npm run package -w docx4j-editor` packages the built extension into `packages/editor-vscode/docx4j-editor-<version>.vsix`
(git-ignored); `npm run vsix` at the root builds everything first. `scripts/package.mjs`:
- packages through `@vscode/vsce` (pinned, a devDependency) with no dependencies. What goes in is the manifest's `files`,
  an allowlist: `vsce` refuses an entry that matches no file, so a missing folder or notice stops it;
- reads the VSIX back and refuses one over 10 MB, one holding a source map or a path outside the allowlist, one lacking a
  bundle or its notices, and one whose version is not the editor's;
- prints the size and the file count.

`--out <folder>` writes the VSIX elsewhere, and `--check <vsix>` checks one made before. `vsce`'s own failures end the
process they run in, so run the script as a process, never import it to package.

CI packages the VSIX in the Node 22 job after the suite, and keeps it as the `vsix` artifact: any commit's can be installed
by hand with `code --install-extension <file>`. A release publishes the public commit's VSIX, never packaging again
(`push-to-marketplace.yml`, to the Visual Studio Marketplace and Open VSX; `RELEASING.md`).

**The README's links** are absolute, to the public repository's `main`: `vsce` rewrites a relative link against the
repository's root, whatever the package's folder. The screenshot, `media/screenshot.png`, is outside the allowlist and
linked at its address there; `npm run screenshot -w docx4j-editor` makes it again through the suite's harness.

## The icon

`media/icon.svg` is the docx4j mark, redrawn as a vector (the portfolio's `brand/docx4j-mark.svg`), and `media/icon.png` its
256-pixel PNG, which the listings take. Both are committed, so the build needs no SVG renderer. To make the PNG again:

```
rsvg-convert -w 256 -h 256 media/icon.svg -o media/icon.png
```

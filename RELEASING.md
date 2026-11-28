# Releasing

Three of the workspace packages are published to npm: `@docx4j/editor-model` (the headless model),
`@docx4j/editor-relay` (the co-editing relay) and `@docx4j/editor-mcp` (the companion, which bundles the built
application). `@docx4j/editor`, the application itself, is private: it is a site, not a package. The VS Code
extension (`packages/editor-vscode`, `plutext.docx4j-editor`) is published to the Visual Studio Marketplace and to
Open VSX. All five share one version.

They are published from GitHub Actions when a GitHub release is published: the npm packages by
`.github/workflows/push-to-npm.yml`, as `@docx4j/core-ts` and `@docx4j/generated-objects-ts` are, and the extension
by `.github/workflows/push-to-marketplace.yml` (see "The VS Code extension" below).

## Publishing target

Publishing uses npm **trusted publishing**: no npm token is stored anywhere. The settings of each of the three
packages on npmjs.com (scope `@docx4j`) name the publisher GitHub Actions and the repository allowed to publish it:
organisation `plutext`, repository `docx4j-editor`, workflow `push-to-npm.yml`, no environment, with **Allow npm
publish** ticked. Without that tick the run gets its OIDC token, then `npm publish` fails with `E403 ... OIDC
permission denied for this action`: fix the setting and re-run the failed job. Renaming the workflow file breaks
publishing until the three settings are changed to match.

A trusted publisher can only be added to a package that exists, so the first version (0.1.0) was published by hand;
see "First release" below.

npm versions cannot be reused once published, even after an unpublish: a problem found after publishing ships as
the next patch version.

Each package ships `dist/`, `README.md`, `LICENSE` and `THIRD-PARTY-NOTICES.txt`, which its `prepack` writes
(`scripts/check-licences.mjs`). The companion's `dist/app` is the built application, without its source maps.

## The VS Code extension

The extension is published through **trusted publishing** on both registries, as the npm packages are: no token is
stored anywhere.

- **The Visual Studio Marketplace.** The publisher `plutext` (marketplace.visualstudio.com/manage) has a policy on
  its Trusted Publishing page: repository owner `plutext`, repository `docx4j-editor`, workflow
  `push-to-marketplace.yml`, no environment. The workflow runs `vsce publish --oidc`. The first successful token
  exchange activates the policy for good.
- **Open VSX** (Cursor, Windsurf, VSCodium, Theia). The namespace `plutext` has a trusted publisher, registered at
  open-vsx.org/user-settings/trusted-publishers, naming the same owner, repository and workflow, with no
  environment. The workflow runs `ovsx publish --trusted-publishing`. `ovsx` is EPL-2.0, so the workflow runs it
  with `npx` at a pinned version, and no package depends on it.

Renaming the workflow file breaks both until each registry's setting is changed to match.

**The file published is the one CI made.** Every `test.yml` run packages the VSIX in its Node 22 job, checks it
(`packages/editor-vscode/scripts/package.mjs`) and keeps it as the artifact `vsix`. On a release,
`push-to-marketplace.yml`:
1. finds a `test.yml` run that passed for the release's commit;
2. downloads its VSIX and checks it again, the tag against its version included;
3. attaches it to the GitHub release;
4. publishes it to the Marketplace and then to Open VSX.

It never packages again, so the VSIX conveyed is built from the source of the public commit it was released from.
GitHub keeps an artifact for 90 days. To release an older commit, run its `test.yml` again first.

Each publish runs whether or not the other one succeeded. Each skips a version its registry already has, so after a
failure, fix the cause and use **Re-run failed jobs**. A Marketplace version, like an npm version, cannot be
published twice, so a problem found afterwards ships as the next patch version. The manifest's `"preview": true`
shows the listing as Preview.

**The first version (0.3.0).** The Marketplace's documentation does not say whether trusted publishing can create an
extension that does not exist yet. If the Marketplace refuses it, upload the VSIX attached to the release by hand on
the publisher's management page (**New extension**, then **Visual Studio Code**). Every later version is the
workflow's. Open VSX's documentation says no more. If it refuses the first version, upload the same file by hand at
open-vsx.org/publish, signed in as an owner of the `plutext` namespace. That needs no token either.

## Dependencies

`package-lock.json` is committed, and CI installs with `npm ci`, so a release is built from exactly the tested
versions. It does not reach consumers, who resolve each `package.json`'s ranges.

A release that needs a change in `@docx4j/core-ts` (or `@docx4j/generated-objects-ts`) waits for that package's
release: publish it first (its own `RELEASING.md`), then raise the range in every package that names it, run the
checks below, and commit the manifests with the lockfile. The editor is never published against an unreleased
engine.

## First release (0.1.0)

The packages already carried 0.1.0. From a clean checkout of the `v0.1.0` tag, whose CI (`test.yml`) had passed:

```bash
npm ci
npm run build && npm test
npm login                                       # an npm account with publish rights on the @docx4j scope
(cd packages/editor-model && npm publish --access public)
(cd packages/editor-relay && npm publish --access public)
(cd packages/editor-mcp && npm publish --access public)
```

Then, on npmjs.com, the trusted publisher was added to each package's settings (as above, with **Allow npm
publish**). The `v0.1.0` tree has no `push-to-npm.yml`, so its GitHub release publishes nothing. Every later
version follows "Steps".

## Steps

```bash
# 1. If the release needs a newer @docx4j/core-ts, release that first and raise the ranges (see "Dependencies").

# 2. Set the version, the next one after the latest on npm (npm view @docx4j/editor-model version), in every
#    package and in their references to one another; then bring the lockfile along. No tag or commit yet.
#    Use npm 11 or later (Node 24 brings it), which wrote the lockfile: npm 10 drops its libc fields.
node scripts/set-version.mjs 0.1.1
npm install

# 3. From a clean node_modules, check against the locked dependencies and inspect each package
rm -rf node_modules packages/*/node_modules packages/*/dist && npm ci
npm run check-licences && npm run locales:check && npm run typecheck && npm run build && npm test
for p in editor-model editor-relay editor-mcp; do (cd packages/$p && npm pack --dry-run); done
npm run package -w docx4j-editor   # the VSIX, packaged and checked (CI's is the one published)

# 4. Commit, tag and push (the tag is v and the version; the workflows check it), then wait for test.yml, whose
#    VSIX the extension's workflow publishes
git commit -am "Version 0.1.1"
git tag -a v0.1.1 -m "Version 0.1.1"
git push origin main v0.1.1
```

5. Publish: create a GitHub release for the tag at https://github.com/plutext/docx4j-editor/releases/new (choose
   the existing tag, title `docx4j editor 0.1.1`, notes for the version, then **Publish release**; a saved draft
   does not publish).

Publishing the release runs `push-to-npm.yml`. It installs with `npm ci` and stops unless the tag equals the
packages' version. Then it runs the licence and message checks, typecheck, build and tests, checks the tree is
unchanged, and packs and publishes the three packages in dependency order, with provenance.

It also runs `push-to-marketplace.yml`, which attaches the VSIX to the release and publishes it to both registries
(see "The VS Code extension"). When it has run, install the extension from each listing and use it: in VS Code from the
Marketplace, and in VSCodium or Cursor from Open VSX. Open a document, edit it, save it, and Invite the agent. A new
version can take some minutes to appear on the Marketplace.

If the workflow fails before anything is published, fix the problem, move the tag, and re-run it from the Actions tab
(or delete and recreate the release). If it fails part way through publishing, or only the publish fails (the E403
above), fix the cause and use **Re-run failed jobs**: a package already published at the version is skipped. The
extension's workflow is re-run the same way.

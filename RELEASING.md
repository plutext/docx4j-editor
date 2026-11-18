# Releasing

Three of the workspace packages are published to npm: `@docx4j/editor-model` (the headless model),
`@docx4j/editor-relay` (the co-editing relay) and `@docx4j/editor-mcp` (the companion, which bundles the built
application). `@docx4j/editor`, the application itself, is private: it is a site, not a package. All four share one
version.

They are published from GitHub Actions (`.github/workflows/push-to-npm.yml`) when a GitHub release is published, as
`@docx4j/core-ts` and `@docx4j/generated-objects-ts` are.

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

# 4. Commit, tag and push (the tag is v and the version; the workflow checks it), then wait for test.yml
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

If the workflow fails before anything is published, fix the problem, move the tag, and re-run it from the Actions tab
(or delete and recreate the release). If it fails part way through publishing, or only the publish fails (the E403
above), fix the cause and use **Re-run failed jobs**: a package already published at the version is skipped.

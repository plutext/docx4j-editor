# Localising the docx4j editor

The editor's interface (its menus, toolbar, dialogs, panels, status messages and the command
palette) is available in several languages. This page is for anyone who wants to add a
language, review a draft, or add a string to the interface. The design is in ED-003 section 12
and the implementation notes in section 11.22.

**The document is never translated.** Its text, its style names (Word's own, `Heading 1`, as the
styles part spells them), its list labels and its field results appear as they are. The
console's scripting interface, its type declarations and the Office JS lint stay in English, as
does the prompt the ask mode sends to the model.

## How it works

- `packages/editor/src/locales/en.json` is the source of truth: a flat JSON object of stable
  keys (`toolbar.bold`, `status.saved.changed`, `dialog.ok`) to English messages. Keys are named
  by place and meaning and are never renamed once shipped, so a translation does not go stale by
  a refactor.
- Every other language is a file beside it named by its BCP 47 tag: `fr.json`, `pt.json`,
  `zh-Hant.json`. A file is found by its name; nothing else needs registering.
- A key a file lacks is shown in English. A file may be incomplete; it may never be wrong.
- The language follows the browser's preferences (`fr-CA` gets `fr`, `zh-TW` and `zh-HK` get
  `zh-Hant`, `pt-PT` gets `pt` until a `pt-PT.json` exists) and can be chosen in Options. Each
  language is loaded when it is chosen, so a file's size costs only its own users.

## The message syntax

Messages use the part of ICU MessageFormat that most translation tools know:

| Form | Example | What to translate |
|---|---|---|
| Argument | `Saved with changes ({how})` | the words; keep `{how}` and move it where the grammar needs it |
| Plural | `{count, plural, one {# paragraph} other {# paragraphs}}` | the words inside each option; keep `count, plural,` and `#` (the number) |
| Select | `{position, select, first {first} other {last}}` | the words inside each option; keep the selector names |
| Exact match | `{count, plural, =0 {none} other {# items}}` | the words |

Each plural must name exactly the categories your language has, no more and no fewer.
`new Intl.PluralRules('<tag>').resolvedOptions().pluralCategories` in a browser console lists
them. For the languages shipped:

| Categories | Languages |
|---|---|
| `other` | Chinese (both), Japanese, Korean, Thai, Malay, Indonesian, Vietnamese |
| `one`, `other` | German, Dutch, Turkish |
| `one`, `many`, `other` | French, Spanish, Italian, Portuguese (`many` is for large round numbers such as a million) |
| `one`, `few`, `many`, `other` | Polish, Ukrainian, Russian |
| `zero`, `one`, `two`, `few`, `many`, `other` | Arabic |

An ASCII apostrophe directly before `{`, `}` or `#` starts quoted text in ICU. Use your
language's typographic apostrophe (`’`) instead; it is never special. The keys
`command.<name>.arg.<arg>` are the exception to all of this: they describe a command's arguments
in the command palette, quote code in braces, and are plain text. Translate their prose and keep
every code fragment as it is.

## Adding a language

1. Copy `en.json` to `<tag>.json` in the same folder.
2. Put `"_status": "draft, not yet reviewed by a native speaker"` first if the file was
   machine-drafted or you want a reviewer; otherwise `"_status": "reviewed by <your name>,
   <date>"`.
3. Translate the values and only the values. Keep product names (Word, OpenDoPE, XPath), code
   (`w:sdtPr`, `od:xpath`), requirement references (`REQ-007`) and key names in shortcuts
   (`Ctrl+K`) as they are. Where Word has a term for something in your language (Bold, Content
   control, Section break), use Word's term: that is what your users know.
4. Run the check from the repository root:

   ```
   npm run locales:check
   ```

   It prints a coverage line per language (`fr: 1080/1080, draft`) and fails on a key English
   lacks, a message that does not parse, an argument English does not have, or a plural with the
   wrong categories. Missing keys are counted, not errors.
5. Open a pull request. A partial file is welcome.

## Reviewing a draft

The shipped files other than English were drafted by machine translation and carry the draft
status; the application says so under the language selector. A review by a native speaker is
the most useful contribution a file can get: correct what reads wrong, keep terms consistent
with Word's, and change `_status` to `"reviewed by <your name>, <date>"`. Reviewing part of a
file is fine; say which part in the pull request.

## Adding a string to the interface (for developers)

Put the key and its English message in `en.json`, look it up with `t('area.key', { args })`
(`const t = useT()` inside a React component, so it re-renders when the language changes), and
run `npm run locales:check`: it fails on a key used and not defined, and on one defined and not
used. Pass an identifier as a string (a number argument is formatted as a number, with grouping).
The other files need nothing: the new key shows in English until a translator adds it.

## The VS Code extension's words

The VS Code extension (`packages/editor-vscode`) has words of its own, outside the application,
in two files of the format VS Code reads (ED-014 section 4.7):

- `package.nls.json`: the manifest's words, which VS Code shows in its own interface: the
  extension's name and description, its commands' titles in the command palette, its settings.
- `l10n/bundle.l10n.json`: the messages the extension shows itself (notifications, input
  boxes). Each key is the English message, as the code passes it to `vscode.l10n.t('…')`;
  translate the value. `node scripts/l10n.mjs`, run in the package, writes it from the sources.

Beside each is a draft per language, `package.nls.<id>.json` and `l10n/bundle.l10n.<id>.json`,
marked `_status` as the application's are.

**Named by VS Code's language id,** which is not always the application's tag:

| Application | VS Code |
|---|---|
| `zh-Hans` | `zh-cn` |
| `zh-Hant` | `zh-tw` |
| `pt` | `pt-br` |

The others are the same. VS Code reads a bundle at its own id exactly, never a shorter one, so a
bundle named `pt` would never be read for `pt-br`.

**What a person sees** is VS Code's language, not the browser's. That is the person's own only
where a VS Code language pack for it is installed, and English otherwise. Microsoft publishes
current packs for Czech, German, Spanish, French, Italian, Japanese, Korean, Polish, Brazilian
Portuguese, Russian, Turkish and both Chinese. Open VSX also lists a community pack for
Ukrainian and one for Arabic, and Microsoft's packs of 2020 for Indonesian and Dutch, whose
translation of today's VS Code is partial. It lists none for Malay, Thai or Vietnamese, so those
drafts are not seen until one exists. A key a draft lacks shows in English.

**The page and VS Code can speak different languages.** The editor inside the extension follows
VS Code's language unless the person chose another in Options, and runs in a language VS Code
has no pack for. So where the page names one of the extension's commands, it names it as VS
Code's palette shows it, through `{command}`, never by translating the name itself.

**The rules** are the application's, and a few more:

- "docx4j editor" is never translated: the product's name stands wherever English has it.
- `{0}`, `{1}` (the bundle's placeholders) are kept, the same ones as English.
- Code in backquotes in the manifest's words (`` `https://api.anthropic.com` ``) is kept as it is.
- The words for the same things are the application's in that language: Commands, Paste as
  Plain Text, Join, Share, Invite, the agent, Ask, the companion.

`npm run locales:check` checks these files too (`node packages/editor-vscode/scripts/l10n.mjs
--check`): the bundle against the sources, every `%key%` of the manifest present, the files
named by VS Code's ids, and in each draft the rules above. It prints each language's coverage.

## Licensing

Contributions, translations included, come in under Apache-2.0 and are distributed with the
editor under `GPL-3.0-or-later`, as `CONTRIBUTING.md` explains; sign off your commits as it
asks, and disclose machine assistance with a trailer as its section on AI-assisted
contributions says.

## Coverage

<!-- locales:coverage -->
| Language | File | Keys | Status |
|---|---|---|---|
| English | `en` | 1687/1687 | source |
| Arabic | `ar` | 1687/1687 | draft |
| German | `de` | 1687/1687 | draft |
| Spanish | `es` | 1687/1687 | draft |
| French | `fr` | 1687/1687 | draft |
| Indonesian | `id` | 1687/1687 | draft |
| Italian | `it` | 1687/1687 | draft |
| Japanese | `ja` | 1687/1687 | draft |
| Korean | `ko` | 1687/1687 | draft |
| Malay | `ms` | 1687/1687 | draft |
| Dutch | `nl` | 1687/1687 | draft |
| Polish | `pl` | 1687/1687 | draft |
| Portuguese | `pt` | 1687/1687 | draft |
| Russian | `ru` | 1687/1687 | draft |
| Thai | `th` | 1687/1687 | draft |
| Turkish | `tr` | 1687/1687 | draft |
| Ukrainian | `uk` | 1687/1687 | draft |
| Vietnamese | `vi` | 1687/1687 | draft |
| Simplified Chinese | `zh-Hans` | 1687/1687 | draft |
| Traditional Chinese | `zh-Hant` | 1687/1687 | draft |
<!-- /locales:coverage -->

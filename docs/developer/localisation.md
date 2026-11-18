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

## Licensing

Contributions, translations included, come in under Apache-2.0 and are distributed with the
editor under `GPL-3.0-or-later`, as `CONTRIBUTING.md` explains; sign off your commits as it
asks, and disclose machine assistance with a trailer as its section on AI-assisted
contributions says.

## Coverage

<!-- locales:coverage -->
| Language | File | Keys | Status |
|---|---|---|---|
| English | `en` | 1625/1625 | source |
| Arabic | `ar` | 1625/1625 | draft |
| German | `de` | 1625/1625 | draft |
| Spanish | `es` | 1625/1625 | draft |
| French | `fr` | 1625/1625 | draft |
| Indonesian | `id` | 1625/1625 | draft |
| Italian | `it` | 1625/1625 | draft |
| Japanese | `ja` | 1625/1625 | draft |
| Korean | `ko` | 1625/1625 | draft |
| Malay | `ms` | 1625/1625 | draft |
| Dutch | `nl` | 1625/1625 | draft |
| Polish | `pl` | 1625/1625 | draft |
| Portuguese | `pt` | 1625/1625 | draft |
| Russian | `ru` | 1625/1625 | draft |
| Thai | `th` | 1625/1625 | draft |
| Turkish | `tr` | 1625/1625 | draft |
| Ukrainian | `uk` | 1625/1625 | draft |
| Vietnamese | `vi` | 1625/1625 | draft |
| Simplified Chinese | `zh-Hans` | 1625/1625 | draft |
| Traditional Chinese | `zh-Hant` | 1625/1625 | draft |
<!-- /locales:coverage -->

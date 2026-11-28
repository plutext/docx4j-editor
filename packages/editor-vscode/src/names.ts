// SPDX-License-Identifier: GPL-3.0-or-later
//
// Names and addresses (ED-014 section 4.3), with no `vscode` import so the unit tests reach them: the Word types, the
// names Save As may write, untitled documents' names, and a link's address as P28 sorts it.

/** The four Word types the editor claims (P4). */
export const WORD_TYPES: readonly string[] = ['.docx', '.dotx', '.docm', '.dotm'];

/** A name's extension, lower-cased, with its dot; `''` for none. */
export function extensionOf(name: string): string {
  const base = baseName(name);
  const dot = base.lastIndexOf('.');
  return dot <= 0 ? '' : base.slice(dot).toLowerCase();
}

/** The last segment of a path, by either separator. */
export function baseName(path: string): string {
  const parts = path.split(/[\\/]/);
  return parts[parts.length - 1] ?? '';
}

export function isWordName(name: string): boolean {
  return WORD_TYPES.includes(extensionOf(name));
}

export function isMarkdownName(name: string): boolean {
  const ext = extensionOf(name);
  return ext === '.md' || ext === '.markdown';
}

/**
 * Whether Save As may write this name (P25): one of the Word types, or `.xml`, which the page writes as Flat OPC. Any other
 * name is refused with nothing written, since a Word document saved as `notes.md` would open in VS Code's text editor.
 */
export function isSaveAsName(name: string): boolean {
  return isWordName(name) || extensionOf(name) === '.xml';
}

/** The Word document a Markdown file opens as (P26): `notes.md` is `notes.docx`. */
export function markdownTarget(name: string): string {
  const base = baseName(name);
  const dot = base.lastIndexOf('.');
  return `${dot <= 0 ? base : base.slice(0, dot)}.docx`;
}

/**
 * The first free untitled name (P17): the name asked for, else the next number. Each page numbers its new documents from
 * one, so a second page's `Document1.docx` becomes `Document2.docx`; a name without a number takes " 2", " 3" and on.
 */
export function untitledName(wanted: string, taken: (name: string) => boolean): string {
  if (!taken(wanted)) return wanted;
  const dot = wanted.lastIndexOf('.');
  const stem = dot <= 0 ? wanted : wanted.slice(0, dot);
  const ext = dot <= 0 ? '' : wanted.slice(dot);
  const numbered = /^(.*?)(\d+)$/.exec(stem);
  const prefix = numbered ? numbered[1]! : `${stem} `;
  let n = numbered ? Number(numbered[2]) + 1 : 2;
  for (;; n++) {
    const name = `${prefix}${n}${ext}`;
    if (!taken(name)) return name;
  }
}

/** A link's address as the extension follows it (P28). */
export type LinkTarget =
  /** `http:`, `https:` and `mailto:`, through `env.openExternal`, which asks before an untrusted site. */
  | { kind: 'external'; url: string }
  /** A `file:` address, normalised to a `file:` URI. */
  | { kind: 'file'; uri: string }
  /** An absolute path with no scheme: a drive path, a UNC path or a POSIX one, `/`-separated. */
  | { kind: 'path'; path: string }
  /** A path relative to the document's folder, its segments joined by `/`. */
  | { kind: 'relative'; path: string }
  /** Any other scheme: a document's links are its author's choice, not the person's. */
  | { kind: 'refused'; scheme: string };

const EXTERNAL = new Set(['http:', 'https:', 'mailto:']);
const SCHEME = /^([A-Za-z][A-Za-z0-9+.-]*):/;
const DRIVE = /^\/*([A-Za-z]:)(\/.*)?$/;

/** Without its fragment and query, which a file has no use for. */
function pathOnly(s: string): string {
  return s.replace(/[?#].*$/, '');
}

function decoded(s: string): string {
  try { return decodeURIComponent(s); } catch { return s; }
}

/**
 * Sorts a link's address (P28). Word writes a relative link with `\` and a `file:` link in several forms (`file:///C:\…`,
 * `file://server/share/…`, `file:C:/…`), so `\` becomes `/` first. A Windows drive path and a UNC path are files too.
 */
export function linkTarget(address: string): LinkTarget {
  const raw = address.trim();
  const scheme = SCHEME.exec(raw)?.[1]?.toLowerCase();
  // A drive letter is not a scheme.
  if (scheme !== undefined && scheme.length > 1) {
    if (EXTERNAL.has(`${scheme}:`)) return { kind: 'external', url: raw };
    if (scheme !== 'file') return { kind: 'refused', scheme: `${scheme}:` };
    const rest = pathOnly(raw.slice(raw.indexOf(':') + 1).replace(/\\/g, '/'));
    const drive = DRIVE.exec(rest);
    if (drive) return { kind: 'file', uri: `file:///${drive[1]!.toUpperCase()}${drive[2] ?? '/'}` };
    // `file://server/share/…`, and the four-slash form some writers give a UNC path, `file:////server/share/…`.
    if (/^\/\/[^/]/.test(rest)) return { kind: 'file', uri: `file:${rest}` };
    if (/^\/{4,}[^/]/.test(rest)) return { kind: 'file', uri: `file://${rest.replace(/^\/+/, '')}` };
    if (rest.startsWith('/')) return { kind: 'file', uri: `file:///${rest.replace(/^\/+/, '')}` };
    return relative(rest);
  }
  const path = pathOnly(raw.replace(/\\/g, '/'));
  if (/^[A-Za-z]:(\/|$)/.test(path)) return { kind: 'path', path: `${path[0]!.toUpperCase()}${path.slice(1)}` };
  if (/^\/\/[^/]/.test(path) || path.startsWith('/')) return { kind: 'path', path };
  return relative(path);
}

function relative(path: string): LinkTarget {
  const segments = decoded(path).split('/').filter((s) => s !== '' && s !== '.');
  if (segments.length === 0) return { kind: 'refused', scheme: '' };
  return { kind: 'relative', path: segments.join('/') };
}

/** A title or category as the extension host hands the manifest over: translated, `{ value, original }`; else a string. */
function wordOf(value: unknown): string | undefined {
  const text = typeof value === 'string' ? value : value && typeof value === 'object' ? (value as { value?: unknown }).value : undefined;
  // An unresolved `%key%`: VS Code running from source leaves the manifest's words so.
  return typeof text === 'string' && text.trim() !== '' && !/^%[^%]+%$/.test(text) ? text : undefined;
}

/**
 * The names VS Code's palette shows for the commands `ids` (ED-014 section 4.7, P49): `<category>: <title>`, from the
 * manifest as the extension reads it (`packageJSON`). VS Code puts a title or category its language translated otherwise
 * than English as `{ value, original }` and leaves any other a string (section 9.25 item 2), so both are read. A command
 * the manifest lacks, or names in words it cannot read, is left out, and the page says it in English.
 */
export function paletteNames(manifest: unknown, ids: readonly string[]): Record<string, string> {
  const commands = (manifest as { contributes?: { commands?: unknown } } | undefined)?.contributes?.commands;
  const out: Record<string, string> = {};
  if (!Array.isArray(commands)) return out;
  for (const c of commands as { command?: unknown; title?: unknown; category?: unknown }[]) {
    if (typeof c?.command !== 'string' || !ids.includes(c.command)) continue;
    const title = wordOf(c.title);
    if (title === undefined) continue;
    const category = wordOf(c.category);
    out[c.command] = category !== undefined ? `${category}: ${title}` : title;
  }
  return out;
}

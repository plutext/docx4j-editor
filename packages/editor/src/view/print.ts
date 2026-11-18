// SPDX-License-Identifier: GPL-3.0-or-later
//
// Printing (ED-006 level 1): the pages. Each section of the document prints on a named page of its own
// size and margins (`@page`), and the blocks of a section are given its page by the `.section-break`
// widget the view draws after each section's last paragraph: a rule per section, `~` over the widget's
// following siblings, in section order, so the nearest section break before a block decides its page
// (section 3.3). No decoration per block: the rules are written at `beforeprint`, from the document as
// it is then, and cost nothing while typing. The section break's own kind is on the section it starts
// (ECMA-376 17.6.22, `w:type`: how the contents of this section are placed relative to the previous).
import type * as wml from '@docx4j/generated-objects-ts/modules/org_docx4j_wml';
import { revive, type EditorDocument } from '@docx4j/editor-model';
import type { Node as PMNode } from 'prosemirror-model';

/** The document settings a page depends on, from the settings part. */
export interface PrintSettings {
  /** `w:mirrorMargins`: the left margin is the inside one, the right the outside, swapped on left-hand pages. */
  mirrorMargins: boolean;
  /** `w:gutterAtTop`: the gutter is added to the top margin rather than the left. */
  gutterAtTop: boolean;
}

export const DEFAULT_PRINT_SETTINGS: PrintSettings = { mirrorMargins: false, gutterAtTop: false };

/** Twips to points, rounded to a hundredth. */
const pt = (twips: number): number => Math.round(twips / 20 * 100) / 100;

/** The page margin kept at the paper's left and right, a strip a printer cannot reach anyway (ED-006 section 9.10). */
const PAPER_EDGE = 18;

interface PageSetup {
  width: number;
  height: number;
  top: number;
  right: number;
  bottom: number;
  left: number;
  gutter: number;
}

/** A section's page in points: Word's defaults (A4, an inch all round) for what it does not say; a negative top or bottom ("exactly") as its size. */
function setupOf(sectPr: wml.SectPr | null | undefined): PageSetup {
  const size = sectPr?.pgSz;
  const m = sectPr?.pgMar;
  const num = (v: unknown, d: number): number => (typeof v === 'number' ? v : typeof v === 'string' && v !== '' ? Number(v) : d);
  return {
    width: pt(num(size?.w, 11906)),
    height: pt(num(size?.h, 16838)),
    top: pt(Math.abs(num(m?.top, 1440))),
    right: pt(num(m?.right, 1440)),
    bottom: pt(Math.abs(num(m?.bottom, 1440))),
    left: pt(num(m?.left, 1440)),
    gutter: pt(num(m?.gutter, 0)),
  };
}

const sameSetup = (a: PageSetup, b: PageSetup): boolean =>
  a.width === b.width && a.height === b.height && a.top === b.top && a.right === b.right && a.bottom === b.bottom && a.left === b.left && a.gutter === b.gutter;

/** How a section starts: its `w:type`, `nextPage` when it says nothing (Word's default). */
function startOf(sectPr: wml.SectPr | null | undefined): string {
  return (sectPr?.type as { val?: string } | undefined)?.val ?? 'nextPage';
}

/** The block root the rules address: the body's ProseMirror element. */
const ROOT = '.docx-editor.body > .ProseMirror';

/**
 * The print rules for the sections, in document order (the paragraph-level ones then the final one,
 * as `sections()` in `css.ts` lists them): an `@page` per distinct page, the blocks' page names, and
 * the breaks between sections. A section of the same page as the one before reuses its name, since
 * CSS breaks the page wherever the name changes: a continuous break then breaks nothing, and a
 * next-page one breaks after the previous section's last paragraph. An odd- or even-page start asks
 * for a right- or left-hand page; browsers print it as a plain page break (ED-006 section 9.1, case F).
 */
export function printPageRules(secs: readonly (wml.SectPr | null | undefined)[], settings: PrintSettings = DEFAULT_PRINT_SETTINGS): string {
  const rules: string[] = [];
  const names: string[] = [];
  const setups = secs.map(setupOf);
  const insideOf = (setup: PageSetup): number => setup.left + (settings.gutterAtTop ? 0 : setup.gutter);
  // Chromium cuts whatever stands outside a page's margins, where Word prints it: a table's edge, which stands a cell's
  // margin out from the text by default, a negative indent, a change bar. So the margins give way to padding, the same
  // amount on every section's page and on both sides, a left-hand page's being the other way round: the text stays put.
  const given = Math.max(0, Math.min(...setups.flatMap((s) => [insideOf(s), s.right])) - PAPER_EDGE);
  const less = (margin: number): number => Math.round((margin - given) * 100) / 100;
  let previous: PageSetup | undefined;
  secs.forEach((_sectPr, i) => {
    const setup = setups[i]!;
    if (previous && sameSetup(previous, setup)) {
      names.push(names[i - 1]!);
    } else {
      const name = `docx-s${i}`;
      names.push(name);
      const top = setup.top + (settings.gutterAtTop ? setup.gutter : 0);
      rules.push(`@page ${name} { size: ${setup.width}pt ${setup.height}pt; margin: ${top}pt ${less(setup.right)}pt ${setup.bottom}pt ${less(insideOf(setup))}pt; }`);
      // Mirror margins: on a left-hand page the inside margin (and its gutter) is on the right.
      if (settings.mirrorMargins) rules.push(`@page ${name}:left { margin-left: ${less(setup.right)}pt; margin-right: ${less(insideOf(setup))}pt; }`);
    }
    previous = setup;
  });
  if (names.length === 0) return '';
  const blocks: string[] = [`${ROOT}, .print-notes { padding-inline: ${Math.round(given * 100) / 100}pt; }`, `${ROOT} > * { page: ${names[0]}; }`];
  for (let i = 1; i < secs.length; i++) {
    // The blocks after section i - 1's break widget belong to section i, unless a later break says otherwise.
    blocks.push(`${ROOT} > .section-break[data-section="${i - 1}"] ~ * { page: ${names[i]}; }`);
    const start = startOf(secs[i]);
    // The break goes after the previous section's last paragraph: the widget may be followed by a page boundary's.
    const last = `${ROOT} > :has(+ .section-break[data-section="${i - 1}"])`;
    if (start === 'oddPage') blocks.push(`${last} { break-after: right; }`);
    else if (start === 'evenPage') blocks.push(`${last} { break-after: left; }`);
    else if (start === 'nextPage' && names[i] === names[i - 1]) blocks.push(`${last} { break-after: page; }`);
  }
  // The notes printed after the body (section 3.6) are on the last section's page, so no break is forced before them.
  blocks.push(`.print-notes { page: ${names[names.length - 1]}; }`);
  return `${rules.join('\n')}\n@media print {\n${blocks.join('\n')}\n}\n`;
}

/**
 * The sections of a view's document in order, the paragraph-level ones then the final one, read from
 * the view's own document (whose `finalSectPr` the view keeps in step with the store), never from
 * `session.doc`, which rebuilds the document from the store.
 */
export function printSections(doc: PMNode): (wml.SectPr | null)[] {
  const out: (wml.SectPr | null)[] = [];
  doc.descendants((node) => {
    if (node.type.name !== 'paragraph') return true;
    const s = node.attrs['sectPr'] as unknown;
    if (s) out.push(revive(s) as wml.SectPr);
    return false;
  });
  const final = doc.attrs['finalSectPr'] as unknown;
  out.push(final ? (revive(final) as wml.SectPr) : null);
  return out;
}

/** The settings part's mirror margins and gutter position. */
export async function printSettingsOf(session: EditorDocument): Promise<PrintSettings> {
  const settings = await session.main.documentSettingsPart?.readContents();
  const on = (v: unknown): boolean => v !== undefined && v !== null && (v as { val?: unknown }).val !== false;
  return { mirrorMargins: on(settings?.mirrorMargins), gutterAtTop: on(settings?.gutterAtTop) };
}

/**
 * A view's print pages: the settings read once and again when a part changes, the rules written into
 * a style element of their own whenever the view is about to print (`beforeprint`, and the Print
 * command before it calls `print`).
 */
export class PrintPages {
  private settings: PrintSettings = DEFAULT_PRINT_SETTINGS;
  private readonly style: HTMLStyleElement;
  private readonly onBeforePrint = (): void => this.prepare();

  constructor(private readonly session: EditorDocument, private readonly doc: () => PMNode, owner: Document) {
    this.style = owner.createElement('style');
    this.style.dataset['docx4jEditor'] = 'print-pages';
    owner.head.appendChild(this.style);
    owner.defaultView?.addEventListener('beforeprint', this.onBeforePrint);
    this.refreshSettings();
  }

  /** The settings part may have changed (a part edit, a peer's). */
  refreshSettings(): void {
    void printSettingsOf(this.session).then((s) => { this.settings = s; }, () => { /* no settings part: the defaults */ });
  }

  /** The rules for the document as it is now. */
  prepare(): void {
    this.style.textContent = printPageRules(printSections(this.doc()), this.settings);
  }

  destroy(): void {
    this.style.ownerDocument.defaultView?.removeEventListener('beforeprint', this.onBeforePrint);
    this.style.remove();
  }
}

// ---------------------------------------------------------------- breaks within the flow (section 3.4)

/** Whether a Word on/off property is stated, and how: undefined when absent, false for `w:val="0"` or `false`. */
function stated(v: unknown): boolean | undefined {
  if (v === undefined || v === null) return undefined;
  const val = (v as { val?: unknown }).val;
  return !(val === false || val === 0 || val === '0' || val === 'false' || val === 'off');
}

/**
 * A paragraph's own pagination properties as the `data-*` attributes print rules read: `1` or `0` where
 * the paragraph's `w:pPr` states one, nothing where it does not (its style's rule then applies).
 */
export function paragraphPagination(pPr: wml.PPr | null | undefined): Record<string, string> {
  const out: Record<string, string> = {};
  const set = (name: string, v: unknown): void => { const s = stated(v); if (s !== undefined) out[name] = s ? '1' : '0'; };
  set('pageBreakBefore', pPr?.pageBreakBefore);
  set('keepNext', pPr?.keepNext);
  set('keepLines', pPr?.keepLines);
  set('widowControl', pPr?.widowControl);
  return out;
}

/** The print declarations a style's effective `w:pPr` makes: only what it turns on, and widow control turned off. */
export function paginationDeclarations(pPr: wml.PPr | null | undefined): string[] {
  const out: string[] = [];
  if (stated(pPr?.pageBreakBefore)) out.push('break-before: page');
  if (stated(pPr?.keepNext)) out.push('break-after: avoid');
  if (stated(pPr?.keepLines)) out.push('break-inside: avoid');
  // Absent, widow control is on in Word (it writes only the off form), as the browsers' orphans and widows of 2 are.
  if (stated(pPr?.widowControl) === false) out.push('orphans: 1', 'widows: 1');
  return out;
}

/** A row's `w:cantSplit` and `w:tblHeader`, from its `trPr` as plain JSON. */
export function rowPagination(trPr: unknown): { cantSplit: boolean; header: boolean } {
  const items = ((trPr as { cnfStyleOrDivIdOrGridBefore?: { name?: { localPart?: string }; value?: unknown }[] } | null)?.cnfStyleOrDivIdOrGridBefore ?? []);
  const on = (local: string): boolean => items.some((it) => it.name?.localPart === local && stated(it.value) !== false);
  return { cantSplit: on('cantSplit'), header: on('tblHeader') };
}

/** How many rows lead a table as header rows, repeated on each printed page as Word repeats them. */
export function leadingHeaderRows(table: PMNode): number {
  let n = 0;
  while (n < table.childCount && rowPagination(table.child(n).attrs['trPr']).header) n++;
  return n;
}

/** The most header rows a printed table repeats: `print.css` has a rule per count up to this. */
export const MAX_PRINT_HEADER_ROWS = 9;

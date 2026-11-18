// SPDX-License-Identifier: GPL-3.0-or-later
//
// Column widths that minimise a table's height (ED-007). The search is pure (sections 3.3 to 3.5):
// over a table whose cells' heights are known at every width that matters, it finds the widths of
// least total height within the table's width. The measurement is the view's (section 3.2): each
// cell's content cloned into a hidden table under the document's stylesheet, its height read at a
// width, and the widths at which it drops a line found by bisection. `fitTableColumns` joins the
// two and runs the model's `table.columnWidths` with the result.
import type { EditorView } from 'prosemirror-view';
import { Selection, TextSelection } from 'prosemirror-state';
import { TableMap } from 'prosemirror-tables';
import { columnWidthsPx, tableAutoFitOf, textWidthTwips, columnWidthsAside, MIN_COLUMN_WIDTH_PX, sessionOf, rowHeightOf, WORD_RUN, breakOffsetsIn } from '@docx4j/editor-model';
import type { Node as PMNode } from 'prosemirror-model';
import { CELL_MIN_WIDTH } from './nodeviews/index.js';
import { scaleOf } from './zoom.js';

// ---------------------------------------------------------------------------------------------
// The pure search
// ---------------------------------------------------------------------------------------------

/** A cell's height at the widths that matter: `steps` sorted by width ascending, each the height from that width up to the next step's. */
export interface CellHeights {
  row: number;
  col: number;
  colspan: number;
  rowspan: number;
  /**
   * `[width, height, breaks?]` triples, widths ascending; the first is the narrowest width measured. `breaks` is how
   * many times a word breaks mid-word at that width (a word in three pieces counts two), present only below the
   * cell's longest word, when the floors were let go (section 8.1 item 12); absent means none.
   */
  steps: [number, number, number?][];
}

export interface FitProblem {
  rows: number;
  cols: number;
  cells: CellHeights[];
  /** Each column's floor, in the same units as the steps. */
  minWidths: number[];
  /** The width the columns must fit in. */
  total: number;
  /** The widths the table has now, for the tie-break and the report. */
  current: number[];
}

export interface FitResult {
  widths: number[];
  height: number;
  heightBefore: number;
  /** The mid-word breaks at the widths chosen and at the current ones. */
  breaks: number;
  breaksBefore: number;
  /** Whether the exhaustive search ran (else greedy with exchange). */
  exhaustive: boolean;
}

/** The limit under which every combination is tried (section 3.4). */
export const EXHAUSTIVE_LIMIT = 50_000;

/** A cell's height at a width: the height of the last step whose width is not above it (the first step's below it). */
export function heightAt(steps: readonly [number, number, number?][], width: number): number {
  let h = steps[0]?.[1] ?? 0;
  for (const [w, sh] of steps) { if (w <= width) h = sh; else break; }
  return h;
}

/** A cell's mid-word breaks at a width, read as `heightAt` reads the height. */
export function breaksAt(steps: readonly [number, number, number?][], width: number): number {
  let b = steps[0]?.[2] ?? 0;
  for (const [w, , sb] of steps) { if (w <= width) b = sb ?? 0; else break; }
  return b;
}

/**
 * What a break outweighs: the search minimises the mid-word breaks first and the height second (Jason, 2026-10-15:
 * as few words broken as possible, and none twice where once would do), so a break counts more than any height.
 */
export const BREAK_WEIGHT = 1_000_000;

/** The table's mid-word breaks at a choice of widths, every cell's summed. */
export function tableBreaks(problem: FitProblem, widths: readonly number[]): number {
  let breaks = 0;
  for (const cell of problem.cells) {
    let w = 0;
    for (let k = 0; k < cell.colspan; k++) w += widths[cell.col + k] ?? 0;
    breaks += breaksAt(cell.steps, w);
  }
  return breaks;
}

/** What the search minimises: the breaks weighted over the height. */
function cost(problem: FitProblem, widths: readonly number[]): number {
  return tableBreaks(problem, widths) * BREAK_WEIGHT + tableHeight(problem, widths);
}

/** The table's height at a choice of widths: each row the tallest of its cells, a spanning cell over its columns' sum and its height shared over its rows. */
export function tableHeight(problem: FitProblem, widths: readonly number[]): number {
  const rows = new Array<number>(problem.rows).fill(0);
  for (const cell of problem.cells) {
    let w = 0;
    for (let k = 0; k < cell.colspan; k++) w += widths[cell.col + k] ?? 0;
    const h = heightAt(cell.steps, w) / cell.rowspan;
    for (let r = cell.row; r < cell.row + cell.rowspan && r < problem.rows; r++) if (h > rows[r]!) rows[r] = h;
  }
  return rows.reduce((a, b) => a + b, 0);
}

/** Each column's candidate widths: its floor and the step widths of the cells that span it alone, sorted, those at or above the floor and within the total. */
export function candidateWidths(problem: FitProblem): number[][] {
  const out: number[][] = [];
  for (let c = 0; c < problem.cols; c++) {
    const set = new Set<number>([problem.minWidths[c] ?? 0]);
    for (const cell of problem.cells) {
      if (cell.col !== c || cell.colspan !== 1) continue;
      for (const [w] of cell.steps) if (w >= (problem.minWidths[c] ?? 0) && w <= problem.total) set.add(w);
    }
    out.push([...set].sort((a, b) => a - b));
  }
  return out;
}

/** Less cost (breaks, then height) first, then less width, then nearer the current widths. */
function better(problem: FitProblem, a: { cost: number; widths: number[] } | undefined, b: { cost: number; widths: number[] }): boolean {
  if (!a) return true;
  if (b.cost !== a.cost) return b.cost < a.cost;
  const sum = (ws: number[]): number => ws.reduce((x, y) => x + y, 0);
  const sa = sum(a.widths), sb = sum(b.widths);
  if (sb !== sa) return sb < sa;
  const dist = (ws: number[]): number => ws.reduce((d, w, i) => d + Math.abs(w - (problem.current[i] ?? w)), 0);
  return dist(b.widths) < dist(a.widths);
}

/** The slack after the chosen widths shared in proportion, so the table keeps its width (section 3.4). */
export function withSlack(widths: readonly number[], total: number): number[] {
  const sum = widths.reduce((a, b) => a + b, 0);
  if (sum <= 0 || sum >= total) return [...widths];
  const factor = total / sum;
  return widths.map((w) => w * factor);
}

/** Every combination of the candidates whose sum fits, the best kept. */
function exhaustive(problem: FitProblem, candidates: number[][]): { widths: number[]; cost: number } | undefined {
  let best: { widths: number[]; cost: number } | undefined;
  const chosen = new Array<number>(problem.cols).fill(0);
  // The least the columns after `c` need, for pruning.
  const restMin: number[] = new Array<number>(problem.cols + 1).fill(0);
  for (let c = problem.cols - 1; c >= 0; c--) restMin[c] = restMin[c + 1]! + candidates[c]![0]!;
  const go = (c: number, used: number): void => {
    if (c === problem.cols) {
      const trial = { widths: [...chosen], cost: cost(problem, chosen) };
      if (better(problem, best, trial)) best = trial;
      return;
    }
    for (const w of candidates[c]!) {
      if (used + w + restMin[c + 1]! > problem.total + 1e-6) break;
      chosen[c] = w;
      go(c + 1, used + w);
    }
  };
  go(0, 0);
  return best;
}

/** Start at the floors; give width to the column whose next candidate saves most cost per pixel; then exchange pairs while it helps. */
function greedy(problem: FitProblem, candidates: number[][]): { widths: number[]; cost: number } | undefined {
  const index = new Array<number>(problem.cols).fill(0);
  const widthsOf = (): number[] => index.map((i, c) => candidates[c]![i]!);
  let widths = widthsOf();
  let used = widths.reduce((a, b) => a + b, 0);
  if (used > problem.total + 1e-6) return undefined;
  let height = cost(problem, widths);
  for (;;) {
    let pick = -1;
    let gain = 0;
    let pickHeight = height;
    for (let c = 0; c < problem.cols; c++) {
      const next = candidates[c]![index[c]! + 1];
      if (next === undefined) continue;
      const extra = next - widths[c]!;
      if (used + extra > problem.total + 1e-6) continue;
      const trial = [...widths]; trial[c] = next;
      const h = cost(problem, trial);
      const g = (height - h) / Math.max(1, extra);
      if (h < height && g > gain) { gain = g; pick = c; pickHeight = h; }
    }
    if (pick < 0) break;
    index[pick]!++;
    widths = widthsOf();
    used = widths.reduce((a, b) => a + b, 0);
    height = pickHeight;
  }
  // Exchange: one column back a step, another forward, kept when the height falls or holds at less width.
  let improved = true;
  while (improved) {
    improved = false;
    for (let a = 0; a < problem.cols && !improved; a++) {
      if (index[a]! === 0) continue;
      for (let b = 0; b < problem.cols && !improved; b++) {
        if (b === a) continue;
        const nextB = candidates[b]![index[b]! + 1];
        if (nextB === undefined) continue;
        const trial = [...widths];
        trial[a] = candidates[a]![index[a]! - 1]!;
        trial[b] = nextB;
        const sum = trial.reduce((x, y) => x + y, 0);
        if (sum > problem.total + 1e-6) continue;
        const h = cost(problem, trial);
        if (h < height || (h === height && sum < used)) {
          index[a]!--; index[b]!++;
          widths = widthsOf(); used = sum; height = h;
          improved = true;
        }
      }
    }
  }
  return { widths, cost: height };
}

/**
 * The widths of fewest mid-word breaks, exactly: a column's breaks are its own single-span cells' at its width, so the
 * choice is a multiple-choice knapsack over the columns within the total, solved by pixel. Of the allocations with the
 * fewest breaks, the narrowest, so the most width is left for the height. The greedy heuristic alone bought the cheapest
 * break removals first and stranded what was left (Jason's table, section 8.1 item 12). Undefined when nothing fits.
 */
export function fewestBreaks(problem: FitProblem, candidates: number[][]): number[] | undefined {
  const budget = Math.floor(problem.total + 1e-6);
  if (budget < 0) return undefined;
  const columnBreaks = (c: number, width: number): number => {
    let b = 0;
    for (const cell of problem.cells) if (cell.colspan === 1 && cell.col === c) b += breaksAt(cell.steps, width);
    return b;
  };
  let dp = new Float64Array(budget + 1).fill(Infinity);
  dp[0] = 0;
  const choices: Int32Array[] = [];
  for (let c = 0; c < problem.cols; c++) {
    const next = new Float64Array(budget + 1).fill(Infinity);
    const choice = new Int32Array(budget + 1).fill(-1);
    for (const [i, w] of candidates[c]!.entries()) {
      const px = Math.ceil(w - 1e-6);
      const b = columnBreaks(c, w);
      for (let t = px; t <= budget; t++) {
        const v = dp[t - px]! + b;
        if (v < next[t]!) { next[t] = v; choice[t] = i; }
      }
    }
    dp = next;
    choices.push(choice);
  }
  let bestT = -1;
  for (let t = 0; t <= budget; t++) if (dp[t]! < Infinity && (bestT < 0 || dp[t]! < dp[bestT]!)) bestT = t;
  if (bestT < 0) return undefined;
  const widths = new Array<number>(problem.cols).fill(0);
  for (let c = problem.cols - 1, t = bestT; c >= 0; c--) {
    const w = candidates[c]![choices[c]![t]!]!;
    widths[c] = w;
    t -= Math.ceil(w - 1e-6);
  }
  return widths;
}

/** The widths of least cost (breaks, then height) within the total (section 3.4), the slack shared; undefined when the floors alone do not fit. */
export function fitWidths(problem: FitProblem): FitResult | undefined {
  let candidates = candidateWidths(problem);
  const combinations = candidates.reduce((n, c) => n * c.length, 1);
  const useExhaustive = combinations <= EXHAUSTIVE_LIMIT;
  if (!useExhaustive && tableBreaks(problem, candidates.map((c) => c[0]!)) > 0) {
    // The fewest breaks first, exactly; the greedy then spends the rest on the height from those widths, never under them.
    const base = fewestBreaks(problem, candidates);
    if (!base) return undefined;
    candidates = candidates.map((list, c) => list.filter((w) => w >= base[c]!));
  }
  const best = useExhaustive ? exhaustive(problem, candidates) : greedy(problem, candidates);
  if (!best) return undefined;
  const widths = withSlack(best.widths, problem.total);
  return {
    widths,
    height: tableHeight(problem, widths),
    heightBefore: tableHeight(problem, problem.current),
    breaks: tableBreaks(problem, widths),
    breaksBefore: tableBreaks(problem, problem.current),
    exhaustive: useExhaustive,
  };
}

/**
 * The widths at which `measure` falls, between `lo` and `hi`, by bisection to the pixel (section 3.3): the height
 * is non-increasing in the width. Returns `[width, height]` steps ascending, the first at `lo`.
 */
export function breakpoints(measure: (width: number) => number, lo: number, hi: number): [number, number][] {
  const out: [number, number][] = [];
  const hLo = measure(lo);
  out.push([lo, hLo]);
  const hHi = hi > lo ? measure(hi) : hLo;
  if (hHi >= hLo) return out;
  // Between two widths of different height, find the narrowest width of the smaller height, then recurse on each side.
  const find = (a: number, ha: number, b: number, hb: number): void => {
    if (hb >= ha || b - a <= 0) return;
    let x = a, y = b, hy = hb;
    while (y - x > 1) {
      const m = Math.floor((x + y) / 2);
      const hm = measure(m);
      if (hm < ha) { y = m; hy = hm; } else x = m;
    }
    // y is the narrowest width whose height is below ha; between it and b there may be more steps.
    out.push([y, hy]);
    if (hy > hb) find(y, hy, b, hb);
  };
  find(lo, hLo, hi, hHi);
  out.sort((p, q) => p[0] - q[0]);
  return out;
}

// ---------------------------------------------------------------------------------------------
// The measurement, in the browser
// ---------------------------------------------------------------------------------------------

export interface FitReport {
  /** The table's height before and after as drawn, in pixels; `after` undefined when nothing was written. */
  before: number;
  after?: number;
  /** The rows whose stated height (exact or minimum) the fit removed, so they take their content's height. */
  rowsReleased: number;
  /** Whether the table was AutoFit to contents, whose widths the fit now states. */
  leftAutoFit: boolean;
  /** `unbroken`: a column narrower than its longest word was widened to it, and the table is no shorter for it. */
  outcome: 'shorter' | 'unbroken' | 'same' | 'tooLarge' | 'tooNarrow' | 'cancelled' | 'noTable';
  /** The columns' longest words together are wider than the table, so the floors were let go (section 8.1 item 12). */
  floorsLetGo?: boolean;
  /** The words broken mid-word in the table as drawn after the fit (counted when the floors were let go). */
  brokenWords?: number;
  /** Of those, the words an optional break could help (a period or a capital inside), as they read, each once (section 8.1 item 13). */
  breakable?: string[];
  /** The table's position in the view's document when the fit ran; absent when none was found. */
  tablePos?: number;
}

export interface FitOptions {
  /** Called with the measurement's progress; the status line. */
  onProgress?: (done: number, total: number) => void;
  /** Cells beyond which the table is refused as too large (section 6). */
  maxCells?: number;
  /** Cells measured per animation frame. */
  chunk?: number;
  /** The table to fit, by its position in the view's document; the table at the selection when absent. */
  tablePos?: number;
}

const MAX_CELLS = 2000;
const CHUNK = 40;

/** The measuring box: a hidden table under the document's stylesheet, in the editor's own column so the fonts and the table's variables apply. */
class Measurer {
  private readonly box: HTMLDivElement;
  private readonly table: HTMLTableElement;
  private readonly col: HTMLTableColElement;
  private readonly cell: HTMLTableCellElement;

  /** The page's zoom over the host (`view/zoom.ts`): the reads are client pixels, the widths given layout pixels. */
  private readonly scale: number;

  constructor(private readonly host: HTMLElement, source: HTMLTableElement) {
    const doc = host.ownerDocument;
    this.scale = scaleOf(host);
    this.box = doc.createElement('div');
    this.box.className = 'tablefit-measure';
    this.box.setAttribute('aria-hidden', 'true');
    this.box.style.cssText = 'position: absolute; inset-inline-start: 0; inset-block-start: 0; visibility: hidden; pointer-events: none; inline-size: 0; block-size: 0; overflow: visible;';
    this.table = doc.createElement('table');
    // The same classes and variables as the table measured, laid out fixed at the width given.
    this.table.className = source.className.replace(/\bautofit(-contents)?\b/g, '');
    this.table.setAttribute('style', source.getAttribute('style') ?? '');
    this.table.style.tableLayout = 'fixed';
    this.table.style.width = 'auto';
    this.table.style.minWidth = '0';
    const colgroup = doc.createElement('colgroup');
    this.col = doc.createElement('col');
    colgroup.appendChild(this.col);
    this.table.appendChild(colgroup);
    const tbody = doc.createElement('tbody');
    const tr = doc.createElement('tr');
    this.cell = doc.createElement('td');
    tr.appendChild(this.cell);
    tbody.appendChild(tr);
    this.table.appendChild(tbody);
    this.box.appendChild(this.table);
    host.appendChild(this.box);
  }

  /** The cell's content installed: the drawn DOM cloned, with the cell's own attributes and inline style. */
  load(cell: HTMLTableCellElement): void {
    this.cell.setAttribute('style', cell.getAttribute('style') ?? '');
    this.cell.className = cell.className;
    this.cell.replaceChildren(...Array.from(cell.childNodes, (n) => n.cloneNode(true)));
  }

  /** The cell's height at a column width, in pixels. */
  heightAt(width: number): number {
    this.col.style.width = `${width}px`;
    this.table.style.width = `${width}px`;
    return this.cell.getBoundingClientRect().height / this.scale;
  }

  /** The cell's mid-word breaks at the width last given to `heightAt`. */
  breaksAt(): number {
    return brokenWordsIn(this.cell).breaks;
  }

  /** The widest the cell's content needs to break no word (its min-content width), in pixels, padding and borders included. */
  minContentWidth(): number {
    this.table.style.width = 'auto';
    this.col.style.width = '';
    const save = this.cell.style.overflowWrap;
    this.cell.style.overflowWrap = 'normal';
    this.table.style.tableLayout = 'auto';
    this.table.style.width = 'min-content';
    const w = this.cell.getBoundingClientRect().width / this.scale;
    this.table.style.tableLayout = 'fixed';
    this.cell.style.overflowWrap = save;
    return Math.ceil(w);
  }

  destroy(): void { this.box.remove(); }
}

/**
 * The words an element draws over more than one line, mid-word: how many words, and how many breaks (a word in three
 * pieces is two). A word is a run of characters of a text node with no blank, hyphen, dash or soft hyphen, since a
 * line may break after those; its pieces are the distinct lines its rectangles lie on.
 */
export function brokenWordsIn(root: Element): { words: number; breaks: number; texts: string[] } {
  const doc = root.ownerDocument;
  const walker = doc.createTreeWalker(root, NodeFilter.SHOW_TEXT);
  let words = 0;
  let breaks = 0;
  const texts = new Set<string>();
  const range = doc.createRange();
  for (let n = walker.nextNode(); n; n = walker.nextNode()) {
    const text = n.nodeValue ?? '';
    for (const m of text.matchAll(WORD_RUN)) {
      range.setStart(n, m.index);
      range.setEnd(n, m.index + m[0].length);
      const tops = new Set<number>();
      for (const r of range.getClientRects()) if (r.width > 1) tops.add(Math.round(r.top));
      if (tops.size > 1) { words++; breaks += tops.size - 1; texts.add(m[0]); }
    }
  }
  return { words, breaks, texts: [...texts] };
}

const nextFrame = (): Promise<void> => new Promise((resolve) => (typeof requestAnimationFrame === 'function' ? requestAnimationFrame(() => resolve()) : setTimeout(resolve, 0)));

/** The table at `pos`, when a table is there. */
function tableNamed(view: EditorView, pos: number): { node: PMNode; pos: number } | undefined {
  const node = pos >= 0 && pos < view.state.doc.content.size ? view.state.doc.nodeAt(pos) : null;
  return node && node.type.spec.tableRole === 'table' ? { node, pos } : undefined;
}

/**
 * Whether the selection is inside the table at `tablePos` (a cell selection's anchor cell counts).
 */
function selectionIn(view: EditorView, tablePos: number): boolean {
  const $from = view.state.selection.$from;
  for (let d = $from.depth; d > 0; d--) if ($from.node(d).type.spec.tableRole === 'table') return $from.before(d) === tablePos;
  return false;
}

/**
 * Runs a table command over the table at `tablePos`: the commands act on the table at the selection (ED-007 section
 * 3.1; nothing outside the view speaks positions), so when the selection is elsewhere it is put in the table's first
 * cell for the command and given back afterwards, mapped through what the command wrote. The two moves are
 * selection-only transactions, no edit and no undo step.
 */
function runInTable(view: EditorView, tablePos: number, command: () => boolean): boolean {
  if (selectionIn(view, tablePos)) return command();
  const before = view.state.selection;
  const inside = Selection.near(view.state.doc.resolve(tablePos + 1), 1);
  view.dispatch(view.state.tr.setSelection(inside));
  const ok = command();
  const back = (): void => {
    const doc = view.state.doc;
    const from = Math.min(before.from, doc.content.size);
    const to = Math.min(before.to, doc.content.size);
    view.dispatch(view.state.tr.setSelection(before instanceof TextSelection ? TextSelection.between(doc.resolve(from), doc.resolve(to)) : Selection.near(doc.resolve(from))));
  };
  back();
  return ok;
}

/**
 * Fits every table between `from` and `to` (a Markdown import's range; ED-007 section 8.1 item 14), each as
 * `fitTableColumns` fits the one at the selection, in document order; a fit cancelled by an edit ends the batch.
 * The widths a fit writes change no position, so the tables found at the start stay where they were.
 */
export async function fitTablesIn(view: EditorView, run: (name: string, args: unknown) => boolean, from: number, to: number, options: FitOptions = {}): Promise<FitReport[]> {
  const tables: number[] = [];
  const doc = view.state.doc;
  doc.nodesBetween(Math.max(0, from), Math.min(to, doc.content.size), (node, pos) => {
    if (node.type.spec.tableRole === 'table') { tables.push(pos); return false; }
    return true;
  });
  const reports: FitReport[] = [];
  for (const tablePos of tables) {
    if ((view as { isDestroyed?: boolean }).isDestroyed) break;
    const report = await fitTableColumns(view, run, { ...options, tablePos });
    reports.push(report);
    if (report.outcome === 'cancelled') break;
  }
  return reports;
}

/** The innermost table around the selection, with its position. */
function tableAt(view: EditorView): { node: PMNode; pos: number } | undefined {
  const $from = view.state.selection.$from;
  for (let d = $from.depth; d >= 0; d--) {
    const n = $from.node(d);
    if (n.type.name === 'table') return { node: n, pos: d === 0 ? 0 : $from.before(d) };
  }
  return undefined;
}

/**
 * Measures the table at the selection and runs `table.columnWidths` with the widths of least height (section 3.7):
 * one undo step. Cancelled, writing nothing, when the document changes under the measurement.
 */
export async function fitTableColumns(view: EditorView, run: (name: string, args: unknown) => boolean, options: FitOptions = {}): Promise<FitReport> {
  const session = sessionOf(view.state);
  const found = options.tablePos !== undefined ? tableNamed(view, options.tablePos) : tableAt(view);
  if (!session || !found) return { before: 0, leftAutoFit: false, outcome: 'noTable', rowsReleased: 0 };
  const { node: table, pos: tablePos } = found;
  const map = TableMap.get(table);
  const cellCount = map.width * map.height;
  if (cellCount > (options.maxCells ?? MAX_CELLS)) return { before: 0, leftAutoFit: false, outcome: 'tooLarge', tablePos, rowsReleased: 0 };
  const dom = view.nodeDOM(tablePos) as HTMLElement | null;
  const tableEl = dom?.querySelector('table') ?? (dom instanceof HTMLTableElement ? dom : null);
  const host = view.dom.parentElement;
  if (!tableEl || !host) return { before: 0, leftAutoFit: false, outcome: 'noTable', tablePos, rowsReleased: 0 };
  const docAtStart = view.state.doc;
  const textPx = textWidthTwips(view.state, tablePos + 1) / 15;
  const current = columnWidthsPx(table, map, textWidthTwips(view.state, tablePos + 1));
  const mode = tableAutoFitOf(table.attrs['tblPr']);
  const pct = (table.attrs['tblPr'] as { tblW?: { type?: string; w?: number } } | null)?.tblW;
  const currentSum = current.reduce((a, b) => a + b, 0);
  // The width to fit in (section 3.5): a fixed table's own, a percentage of the text width, else the text width; never past the margin
  // unless the table is already, in which case it keeps its width.
  let total = mode === 'fixed' && pct?.type === 'dxa' && pct.w ? pct.w / 15 : mode === 'window' ? ((pct?.w ?? 5000) / 5000) * textPx : textPx;
  if (currentSum > textPx) total = currentSum;
  const leftAutoFit = mode === 'contents' && columnWidthsAside(table);
  // The cells: one per grid cell, spans once.
  const cells: CellHeights[] = [];
  // The rows whose stated height the fit releases (section 3.5): exact or minimum, every one.
  let rowsWithHeight = 0;
  table.forEach((row) => { if (rowHeightOf(row.attrs['trPr'])) rowsWithHeight++; });
  const hardMin = Math.max(CELL_MIN_WIDTH, MIN_COLUMN_WIDTH_PX);
  const minWidths = new Array<number>(map.width).fill(hardMin);
  const measurer = new Measurer(host, tableEl);
  const seen = new Set<number>();
  const positions: { r: number; c: number; pos: number }[] = [];
  for (let r = 0; r < map.height; r++) {
    for (let c = 0; c < map.width; c++) {
      const pos = map.map[r * map.width + c]!;
      if (seen.has(pos)) continue;
      seen.add(pos);
      positions.push({ r, c, pos });
    }
  }
  // The floors first (section 3.5): the drag's, raised to each cell's widest word (over the cell's columns together for a
  // span). When the columns' floors together are wider than the table (section 8.1 item 12), no widths can keep every word
  // whole: the floors are let go, the cells measured down to the drag's minimum with the words breaking as the browser
  // breaks them, and the search minimises the height as before, so a word breaks only where that costs less height than
  // the width it would take.
  const minContents: number[] = [];
  let wordsBreak = false;
  try {
    let done = 0;
    const step = async (): Promise<boolean> => {
      done++;
      if (done % (options.chunk ?? CHUNK) === 0) { options.onProgress?.(done, positions.length * 2); await nextFrame(); }
      return view.state.doc === docAtStart;
    };
    for (const [i, { c, pos }] of positions.entries()) {
      const cellDom = view.nodeDOM(tablePos + 1 + pos) as HTMLTableCellElement | null;
      if (!cellDom) continue;
      measurer.load(cellDom);
      const minContent = measurer.minContentWidth();
      minContents[i] = minContent;
      const colspan = table.nodeAt(pos)!.attrs['colspan'] as number;
      if (colspan === 1 && minContent > minWidths[c]!) minWidths[c] = Math.min(minContent, total);
      if (!(await step())) return { before: 0, leftAutoFit, outcome: 'cancelled', tablePos, rowsReleased: 0 };
    }
    if (minWidths.reduce((a, b) => a + b, 0) > total + 0.5) {
      wordsBreak = true;
      minWidths.fill(hardMin);
    }
    for (const [i, { r, c, pos }] of positions.entries()) {
      const cell = table.nodeAt(pos)!;
      const colspan = cell.attrs['colspan'] as number;
      const rowspan = cell.attrs['rowspan'] as number;
      const cellDom = view.nodeDOM(tablePos + 1 + pos) as HTMLTableCellElement | null;
      if (!cellDom) continue;
      measurer.load(cellDom);
      const minContent = minContents[i] ?? hardMin;
      const lo = colspan === 1 ? minWidths[c]! : Math.max(CELL_MIN_WIDTH, wordsBreak ? CELL_MIN_WIDTH : minContent);
      const hi = Math.max(lo, Math.min(total, Math.floor(total)));
      const steps: [number, number, number?][] = breakpoints((w) => measurer.heightAt(w), Math.floor(lo), hi);
      // A cell narrower now than its floor (a file fit under other fonts, a drag) is as tall as its broken words make it:
      // measured at its width, so the table's height before is the real one and the floor's widths count as an improvement.
      let now = 0;
      for (let k = 0; k < colspan; k++) now += current[c + k] ?? 0;
      if (now < Math.floor(lo)) {
        const broken = measurer.heightAt(now);
        if (broken > (steps[0]?.[1] ?? 0)) steps.unshift([now, broken]);
      }
      // With the floors let go, the steps under this cell's longest word carry their mid-word breaks, which the search
      // minimises before the height.
      if (wordsBreak) {
        for (const step of steps) {
          if (step[0] >= minContent) break;
          measurer.heightAt(step[0]);
          step[2] = measurer.breaksAt();
        }
      }
      cells.push({ row: r, col: c, colspan, rowspan, steps });
      if (!(await step())) return { before: 0, leftAutoFit, outcome: 'cancelled', tablePos, rowsReleased: 0 };
    }
  } finally {
    measurer.destroy();
  }
  if (view.state.doc !== docAtStart) return { before: 0, leftAutoFit, outcome: 'cancelled', tablePos, rowsReleased: 0 };
  // A single-span cell's floor may have risen after its steps were measured: the candidates honour the floors.
  const problem: FitProblem = { rows: map.height, cols: map.width, cells, minWidths, total, current };
  const result = fitWidths(problem);
  // The report's heights are the table's as drawn, read after the view has applied the transaction (its DOM follows
  // the dispatch), so a released row height counts as the widths do.
  const drawnHeight = (): number => {
    const el = view.nodeDOM(tablePos) as HTMLElement | null;
    const t = el?.querySelector('table') ?? (el instanceof HTMLTableElement ? el : null);
    return (t?.getBoundingClientRect().height ?? 0) / scaleOf(host);
  };
  const before = drawnHeight();
  if (!result) return { before, leftAutoFit, outcome: 'tooNarrow', tablePos, rowsReleased: 0 };
  // A column under its floor (a file fit under other fonts, a drag) breaks a word at its edge, which no one wants from
  // the button (section 3.5): the floor's widths are applied even where the table is no shorter for them.
  const belowFloor = current.some((w, i) => w < minWidths[i]! - 0.5);
  const widthsSame = result.widths.every((w, i) => Math.abs(w - current[i]!) < 0.5);
  const unchanged = widthsSame || (!belowFloor && result.height >= result.heightBefore && result.breaks >= result.breaksBefore);
  const broken = (): { brokenWords: number; breakable: string[] } => {
    if (!wordsBreak) return { brokenWords: 0, breakable: [] };
    const el = view.nodeDOM(tablePos) as HTMLElement | null;
    const t = el?.querySelector('table') ?? (el instanceof HTMLTableElement ? el : null);
    const found = t ? brokenWordsIn(t) : { words: 0, texts: [] };
    return { brokenWords: found.words, breakable: found.texts.filter((w) => breakOffsetsIn(w).length > 0) };
  };
  if (unchanged && rowsWithHeight === 0) return { before, leftAutoFit, outcome: 'same', tablePos, rowsReleased: 0, floorsLetGo: wordsBreak, ...broken() };
  const widths = (unchanged ? current : result.widths).map((w) => (w * 15) / 20);
  const ok = runInTable(view, tablePos, () => run('table.columnWidths', { widths, rowHeights: 'content' }));
  if (!ok) return { before, leftAutoFit, outcome: 'same', tablePos, rowsReleased: 0, floorsLetGo: wordsBreak, ...broken() };
  const after = drawnHeight();
  const outcome = after < before - 0.5 ? 'shorter' : !unchanged && belowFloor ? 'unbroken' : 'same';
  return { before, after, leftAutoFit, outcome, tablePos, rowsReleased: rowsWithHeight, floorsLetGo: wordsBreak, ...broken() };
}

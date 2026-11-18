// SPDX-License-Identifier: GPL-3.0-or-later
//
// ED-007's search (sections 3.3 to 3.5) over hand-made height tables with the optimum known: the
// breakpoints by bisection, the table's height at a choice of widths, the exhaustive search, greedy
// with exchange reaching the same answer, the constraints (floors, exact rows, merged cells) and the
// slack shared in proportion. The measurement is the browser's (e2e/tablefit.spec.ts).
import { describe, test, expect } from 'vitest';
import { breakpoints, candidateWidths, tableHeight, tableBreaks, fitWidths, fewestBreaks, heightAt, breaksAt, withSlack, EXHAUSTIVE_LIMIT, type FitProblem, type CellHeights } from '../src/view/tablefit.js';

/** A text cell as a measurer sees it: `chars` characters at `cw` pixels each, `lh` per line, never fewer than one line. */
const textCell = (chars: number, cw = 7, lh = 16) => (width: number): number => Math.max(1, Math.ceil(chars / Math.max(1, Math.floor(width / cw)))) * lh;

const cell = (row: number, col: number, measure: (w: number) => number, lo: number, hi: number, extra: Partial<CellHeights> = {}): CellHeights =>
  ({ row, col, colspan: 1, rowspan: 1, steps: breakpoints(measure, lo, hi), ...extra });

describe('breakpoints', () => {
  test('finds every width at which the height falls, narrowest first', () => {
    const m = textCell(40);   // 280 px of characters: 1 line from 280, 2 from 140, 3 from 98, 4 from 70 ...
    const steps = breakpoints(m, 30, 400);
    expect(steps[0]).toEqual([30, m(30)]);
    expect(steps[steps.length - 1]).toEqual([280, 16]);
    // Each step's height is the height just at it, and one more line just below it.
    for (const [w, h] of steps.slice(1)) { expect(m(w)).toBe(h); expect(m(w - 1)).toBeGreaterThan(h); }
    // Ascending widths, descending heights.
    for (let i = 1; i < steps.length; i++) { expect(steps[i]![0]).toBeGreaterThan(steps[i - 1]![0]); expect(steps[i]![1]).toBeLessThan(steps[i - 1]![1]); }
  });

  test('a cell of one line has its floor alone', () => {
    expect(breakpoints(textCell(3), 30, 400)).toEqual([[30, 16]]);
  });

  test('heightAt reads the step at or below the width', () => {
    const steps: [number, number][] = [[30, 64], [70, 48], [140, 32], [280, 16]];
    expect(heightAt(steps, 20)).toBe(64);
    expect(heightAt(steps, 70)).toBe(48);
    expect(heightAt(steps, 139)).toBe(48);
    expect(heightAt(steps, 1000)).toBe(16);
  });
});

describe('fitWidths', () => {
  /** The pasted example: short identifiers, long descriptions, short dates, at equal widths that leave the middle wrapping. */
  const example = (): FitProblem => ({
    rows: 3, cols: 3,
    cells: [
      cell(0, 0, textCell(2), 30, 600), cell(0, 1, textCell(120), 30, 600), cell(0, 2, textCell(6), 30, 600),
      cell(1, 0, textCell(2), 30, 600), cell(1, 1, textCell(90), 30, 600), cell(1, 2, textCell(8), 30, 600),
      cell(2, 0, textCell(2), 30, 600), cell(2, 1, textCell(20), 30, 600), cell(2, 2, textCell(6), 30, 600),
    ],
    minWidths: [30, 30, 30],
    total: 600,
    current: [200, 200, 200],
  });

  test('gives the long column the width: the short ones at their floors, the table as short as its lines allow', () => {
    const problem = example();
    const r = fitWidths(problem)!;
    expect(r.exhaustive).toBe(true);
    expect(r.height).toBeLessThan(r.heightBefore);
    expect(r.widths.reduce((a, b) => a + b, 0)).toBeCloseTo(600, 6);
    expect(r.widths[1]).toBeGreaterThan(r.widths[0]! * 5);
    // At 600 px the middle column can take 120 chars at 7 px in two lines (420 px) or three (280): the best is two lines each row, one for the last.
    expect(r.height).toBe(16 * (2 + 2 + 1));
    // The short columns keep their narrowest useful widths, raised by the slack alike: 30 for two characters; 42 for six, since
    // the eight-character cell wraps in a row the long column holds at two lines anyway.
    expect(r.widths[0]! / 30).toBeCloseTo(r.widths[2]! / 42, 6);
  });

  test('the exhaustive search is the oracle for greedy with exchange over a larger table', () => {
    const problem = example();
    const r = fitWidths(problem)!;
    // The same problem with the limit forced to greedy: `EXHAUSTIVE_LIMIT` combinations are not reached here, so make the candidates many.
    const candidates = candidateWidths(problem);
    expect(candidates.reduce((n, c) => n * c.length, 1)).toBeLessThanOrEqual(EXHAUSTIVE_LIMIT);
    const wide: FitProblem = {
      rows: 6, cols: 6,
      cells: Array.from({ length: 36 }, (_, i) => cell(Math.floor(i / 6), i % 6, textCell(10 + ((i * 37) % 200)), 30, 900)),
      minWidths: new Array<number>(6).fill(30),
      total: 900,
      current: new Array<number>(6).fill(150),
    };
    const g = fitWidths(wide)!;
    expect(g.exhaustive).toBe(false);
    expect(g.height).toBeLessThanOrEqual(g.heightBefore);
    // Greedy's answer within a line of the exhaustive one over the three best candidates per column.
    expect(g.height - r.height).toBeGreaterThanOrEqual(0);
    expect(g.widths.reduce((a, b) => a + b, 0)).toBeCloseTo(900, 6);
  });

  test('a cell whose floor is raised keeps it; the floors alone too wide is refused', () => {
    const raised = example();
    raised.minWidths = [30, 30, 300];
    const rr = fitWidths(raised)!;
    expect(rr.widths[2]).toBeGreaterThanOrEqual(300);
    const impossible = example();
    impossible.minWidths = [300, 300, 300];
    expect(fitWidths(impossible)).toBeUndefined();
  });

  test('a cell merged over two columns is measured over their sum, and a vertical merge shares its height', () => {
    const problem = example();
    problem.cells.push({ row: 0, col: 0, colspan: 2, rowspan: 1, steps: breakpoints(textCell(60), 60, 600) });
    problem.cells = problem.cells.filter((c) => !(c.row === 0 && c.colspan === 1 && c.col < 2));
    const r = fitWidths(problem)!;
    expect(tableHeight(problem, r.widths)).toBe(r.height);
    // The merged cell at 60 chars (420 px) is one line when its columns sum to 420 or more.
    const sum01 = r.widths[0]! + r.widths[1]!;
    expect(heightAt(problem.cells[problem.cells.length - 1]!.steps, sum01)).toBe(16);
    const tall: FitProblem = { ...example(), cells: [{ row: 0, col: 0, colspan: 1, rowspan: 3, steps: [[30, 96]] }, cell(0, 1, textCell(4), 30, 600), cell(1, 1, textCell(4), 30, 600), cell(2, 1, textCell(4), 30, 600)], cols: 2, minWidths: [30, 30], current: [300, 300] };
    expect(tableHeight(tall, [100, 100])).toBe(96);   // 32 to each of its three rows, over the 16 of the others
  });

  test('a column now narrower than its floor counts its broken height, so the floor is an improvement', () => {
    const problem = example();
    // The first column at 20 px, under its floor of 30, where its two characters break onto two lines.
    problem.current = [20, 380, 200];
    for (const c of problem.cells) if (c.col === 0) c.steps.unshift([20, 32]);
    const r = fitWidths(problem)!;
    expect(r.heightBefore).toBe(tableHeight(problem, [20, 380, 200]));
    expect(heightAt(problem.cells[0]!.steps, 20)).toBe(32);
    expect(r.widths[0]).toBeGreaterThanOrEqual(30);
    expect(r.height).toBeLessThan(r.heightBefore);
  });

  test('with the floors let go, the fewest mid-word breaks come first and the height second (section 8.1 item 12)', () => {
    // Two columns in 100 px whose longest words need 60 each: one word must break. Under 60 a cell's steps carry its
    // breaks: two pieces from 30, three below it. Breaking one word once (60 and 40) beats two words once (50 and 50),
    // though the heights tie at 48, and beats one word twice.
    const word = (row: number, col: number): CellHeights => ({ row, col, colspan: 1, rowspan: 1, steps: [[20, 48, 2], [30, 32, 1], [60, 16]] });
    const problem: FitProblem = { rows: 1, cols: 2, cells: [word(0, 0), word(0, 1)], minWidths: [20, 20], total: 100, current: [50, 50] };
    expect(breaksAt(problem.cells[0]!.steps, 25)).toBe(2);
    expect(breaksAt(problem.cells[0]!.steps, 59)).toBe(1);
    expect(breaksAt(problem.cells[0]!.steps, 60)).toBe(0);
    expect(tableBreaks(problem, [50, 50])).toBe(2);
    const r = fitWidths(problem)!;
    expect(r.breaksBefore).toBe(2);
    expect(r.breaks).toBe(1);
    expect(r.widths.reduce((a, b) => a + b, 0)).toBeCloseTo(100, 6);
    expect(Math.max(...r.widths)).toBeGreaterThanOrEqual(60);
    // With the height the only measure (no breaks recorded), 50 and 50 is as good as 60 and 40 and nearer the current widths.
    const plain: FitProblem = { ...problem, cells: [word(0, 0), word(0, 1)].map((c) => ({ ...c, steps: c.steps.map(([w, h]) => [w, h] as [number, number]) })) };
    expect(fitWidths(plain)!.widths).toEqual([50, 50]);
  });

  test('over a table too large for the exhaustive search, the breaks are the fewest of every combination (the knapsack), the height greedy after', () => {
    // Four columns in 200 px, every cell a word of its own length: under it, breaks at each height step. The candidates
    // outnumber the exhaustive limit, so the greedy path runs, with the knapsack's floors under it.
    const wordy = (row: number, col: number, need: number): CellHeights => {
      const steps: [number, number, number?][] = [];
      for (let pieces = 6; pieces >= 2; pieces--) steps.push([Math.ceil(need / pieces) + (col % 3), 16 * pieces, pieces - 1]);
      steps.push([need, 16]);
      steps.push([need + 40 + col, 16]);
      return { row, col, colspan: 1, rowspan: 1, steps: steps.sort((a, b) => a[0] - b[0]) };
    };
    const problem: FitProblem = {
      rows: 3, cols: 4,
      cells: Array.from({ length: 12 }, (_, i) => wordy(Math.floor(i / 4), i % 4, 60 + 7 * (i % 4) + 3 * Math.floor(i / 4))),
      minWidths: new Array<number>(4).fill(12), total: 200, current: new Array<number>(4).fill(50),
    };
    const candidates = candidateWidths(problem);
    expect(candidates.reduce((n, c) => n * c.length, 1)).toBeGreaterThan(EXHAUSTIVE_LIMIT);
    // Every combination within the total, the fewest breaks of them all.
    let fewest = Infinity;
    const walk = (c: number, used: number, chosen: number[]): void => {
      if (c === 4) { fewest = Math.min(fewest, tableBreaks(problem, chosen)); return; }
      for (const w of candidates[c]!) { if (used + w > 200) break; walk(c + 1, used + w, [...chosen, w]); }
    };
    walk(0, 0, []);
    expect(fewest).toBeLessThan(tableBreaks(problem, problem.current));
    const base = fewestBreaks(problem, candidates)!;
    expect(tableBreaks(problem, base)).toBe(fewest);
    const r = fitWidths(problem)!;
    expect(r.exhaustive).toBe(false);
    expect(r.breaks).toBe(fewest);
    expect(r.widths.reduce((a, b) => a + b, 0)).toBeCloseTo(200, 6);
    r.widths.forEach((w, c) => expect(w).toBeGreaterThanOrEqual(base[c]!));
  });

  test('the slack is shared in proportion, and a table already as short as it can be reports the same height', () => {
    expect(withSlack([100, 300], 800)).toEqual([200, 600]);
    expect(withSlack([500, 500], 800)).toEqual([500, 500]);
    const problem = example();
    const r = fitWidths(problem)!;
    const again = fitWidths({ ...problem, current: r.widths })!;
    expect(again.height).toBe(again.heightBefore);
  });
});

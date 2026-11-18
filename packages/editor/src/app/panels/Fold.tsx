// SPDX-License-Identifier: GPL-3.0-or-later
//
// A panel section that folds (ED-003 section 8.6): a `<details>` whose open state this browser
// remembers per section, so the console folds under reveal codes, the ask box under the
// console, and a Template tab's groups under each other. The heading sits in the summary; a
// top-level section keeps its region role and name, so the panels are found as before.
import { useState, type ReactNode } from 'react';

const KEY = 'docx4j-editor.fold.';

/** Whether a section is open in this browser: what was saved, else the fallback. */
export function loadFold(id: string, fallback = true): boolean {
  try {
    const v = localStorage.getItem(KEY + id);
    return v === null ? fallback : v === 'open';
  } catch { return fallback; }
}

export function saveFold(id: string, open: boolean): void {
  try { localStorage.setItem(KEY + id, open ? 'open' : 'closed'); } catch { /* storage may be unavailable */ }
}

export interface FoldProps {
  /** The storage key: `console`, `ask`, `control.binding`. */
  id: string;
  /** The heading's text; `meta` is shown beside it, inside the summary. */
  title: ReactNode;
  meta?: ReactNode;
  /** The heading level: 2 for a tab's section (a labelled region), 3 for a group inside one (a plain division). */
  level?: 2 | 3;
  className?: string;
  /** The region's name; a level-3 fold with a label is a region too (the ask box). */
  label?: string;
  testId?: string;
  defaultOpen?: boolean;
  children: ReactNode;
}

export function Fold({ id, title, meta, level = 2, className, label, testId, defaultOpen = true, children }: FoldProps) {
  const [open, setOpen] = useState(() => loadFold(id, defaultOpen));
  const onToggle = (e: React.SyntheticEvent<HTMLDetailsElement>): void => {
    const next = e.currentTarget.open;
    if (next === open) return;
    setOpen(next);
    saveFold(id, next);
  };
  const heading = <>{title}{meta !== undefined && meta !== null && meta !== false ? <span className="meta"> {meta}</span> : null}</>;
  const body = (
    <details open={open} onToggle={onToggle} data-fold={id}>
      <summary>{level === 2 ? <h2>{heading}</h2> : <h3>{heading}</h3>}</summary>
      {children}
    </details>
  );
  const cls = `fold${className ? ` ${className}` : ''}`;
  const name = label ?? (level === 2 && typeof title === 'string' ? title : undefined);
  return name !== undefined
    ? <section className={cls} aria-label={name} data-testid={testId}>{body}</section>
    : <div className={cls} data-testid={testId}>{body}</div>;
}

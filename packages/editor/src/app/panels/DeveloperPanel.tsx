// SPDX-License-Identifier: GPL-3.0-or-later
//
// The side panel (ED-002 section 5.4; reworked 2026-09-24 at Jason's request; ED-003 section
// 4.9): one column with tabs like a browser's developer tools, the tabs being what the current
// perspective contributes (presets.tsx): six for Developer, the Template ones for Template, none
// for Edit. Every tab stays mounted once it has been opened, so the console keeps its script and
// the parts list its selection when another tab is in front; the tab strip is a `tablist` the
// arrow keys walk. Beside it, a draggable separator sets the column's width.
import { useCallback, useEffect, useRef, useState } from 'react';
import { PANEL_MIN_WIDTH, type TabId, type TabSpec, type TabContext, type PanelState } from '../perspectives.js';
import { useT, isRightToLeft } from '../i18n.js';
import { tabLabel } from '../labels.js';

export interface DeveloperPanelProps {
  tabs: readonly TabSpec[];
  panel: PanelState;
  onTab: (tab: TabId) => void;
  context: TabContext;
}

export function DeveloperPanel({ tabs, panel, onTab, context }: DeveloperPanelProps) {
  const t = useT();
  const [mounted, setMounted] = useState<Set<TabId>>(() => new Set([panel.tab]));
  useEffect(() => {
    setMounted((m) => (m.has(panel.tab) ? m : new Set([...m, panel.tab])));
  }, [panel.tab]);

  const onKey = (e: React.KeyboardEvent): void => {
    if (tabs.length === 0) return;
    const i = tabs.findIndex((tab) => tab.id === panel.tab);
    // The arrow that points along the reading direction goes to the next tab: Left in a right-to-left chrome.
    const next = isRightToLeft(e.currentTarget) ? 'ArrowLeft' : 'ArrowRight';
    const previous = next === 'ArrowLeft' ? 'ArrowRight' : 'ArrowLeft';
    if (e.key === next) onTab(tabs[(i + 1) % tabs.length]!.id);
    else if (e.key === previous) onTab(tabs[(i - 1 + tabs.length) % tabs.length]!.id);
    else if (e.key === 'Home') onTab(tabs[0]!.id);
    else if (e.key === 'End') onTab(tabs[tabs.length - 1]!.id);
    else return;
    e.preventDefault();
  };

  const show = (tab: TabId): boolean => mounted.has(tab) || tab === panel.tab;

  return (
    <aside className="panels" aria-label={t('panels.aside')} data-testid="panels">
      <div className="tab-strip" role="tablist" aria-label={t('panels.tabs')} onKeyDown={onKey}>
        {tabs.map((tab) => (
          <button key={tab.id} type="button" role="tab" id={`tab-${tab.id}`} aria-selected={tab.id === panel.tab} aria-controls={`tabpanel-${tab.id}`}
            tabIndex={tab.id === panel.tab ? 0 : -1} data-testid={`tab-${tab.id}`} className={tab.id === panel.tab ? 'selected' : ''} onClick={() => onTab(tab.id)}>{tabLabel(tab)}</button>
        ))}
      </div>
      {tabs.map((tab) => show(tab.id) && (
        <div key={tab.id} role="tabpanel" id={`tabpanel-${tab.id}`} aria-labelledby={`tab-${tab.id}`} className="tab-content" hidden={panel.tab !== tab.id}>
          {tab.render(context)}
        </div>
      ))}
    </aside>
  );
}

export interface SplitterProps {
  width: number;
  onWidth: (width: number) => void;
  /** The element whose right edge the panel ends at (the workspace), for the pointer arithmetic. */
  container: HTMLElement | null;
}

/** The separator between the page area and the panel: drag it, or use the arrow keys, to resize the panel. */
export function Splitter({ width, onWidth, container }: SplitterProps) {
  const t = useT();
  const dragging = useRef(false);
  /** Where in the separator the pointer went down, so the panel does not jump by that much at the first move. */
  const offset = useRef(0);
  const clamp = useCallback((w: number): number => {
    const max = container ? Math.max(PANEL_MIN_WIDTH, container.clientWidth - 320) : Number.POSITIVE_INFINITY;
    return Math.min(max, Math.max(PANEL_MIN_WIDTH, w));
  }, [container]);

  // The panel is the workspace's last column: at its right edge in a left-to-right chrome, at its left edge in a right-to-left one (the grid follows `dir`).
  const widthAt = (clientX: number, rtl: boolean): number => {
    const r = container!.getBoundingClientRect();
    return rtl ? clientX - r.left : r.right - clientX;
  };
  const onPointerDown = (e: React.PointerEvent<HTMLDivElement>): void => {
    dragging.current = true;
    offset.current = container ? widthAt(e.clientX, isRightToLeft(container)) - width : 0;
    e.currentTarget.setPointerCapture(e.pointerId);
    e.preventDefault();
  };
  const onPointerMove = (e: React.PointerEvent<HTMLDivElement>): void => {
    if (!dragging.current || !container) return;
    onWidth(clamp(widthAt(e.clientX, isRightToLeft(container)) - offset.current));
  };
  const onPointerUp = (e: React.PointerEvent<HTMLDivElement>): void => {
    dragging.current = false;
    e.currentTarget.releasePointerCapture(e.pointerId);
  };
  const onKey = (e: React.KeyboardEvent): void => {
    const step = e.shiftKey ? 80 : 16;
    // The arrow that moves the separator towards the page widens the panel: Left when the panel is on the right, Right when it is on the left.
    const wider = isRightToLeft(e.currentTarget) ? 'ArrowRight' : 'ArrowLeft';
    if (e.key === wider) onWidth(clamp(width + step));
    else if (e.key === 'ArrowLeft' || e.key === 'ArrowRight') onWidth(clamp(width - step));
    else return;
    e.preventDefault();
  };

  return (
    <div className="splitter" role="separator" aria-orientation="vertical" aria-label={t('panels.splitter')} aria-valuenow={Math.round(width)} aria-valuemin={PANEL_MIN_WIDTH}
      tabIndex={0} data-testid="splitter" title={t('panels.splitter.title')}
      onPointerDown={onPointerDown} onPointerMove={onPointerMove} onPointerUp={onPointerUp} onPointerCancel={onPointerUp} onKeyDown={onKey} />
  );
}

// SPDX-License-Identifier: GPL-3.0-or-later
//
// A bottom sheet (ED-013 section 3.6): the compact layout's menu, its changes (Review) and its comments. It is a modal
// dialog as `Dialog` is: labelled, Escape closing it, Tab kept inside it, the focus given back where it was when it
// closes (ED-005 section 12.41). It covers up to 85% of the height and scrolls inside; a swipe down on its handle or
// its title (the body scrolls, so a swipe there is the list's), its close button or a tap outside it close it.
import { useEffect, useRef, useState, type KeyboardEvent, type PointerEvent, type ReactNode } from 'react';
import { FOCUSABLE, focusIsNowhere } from './Dialogs.js';
import { useT } from './i18n.js';

/** How far down the handle is dragged, in CSS px, before the sheet closes. */
const SWIPE_CLOSE = 60;

export function Sheet({ title, onClose, children, testId, className }: { title: string; onClose: () => void; children: ReactNode; testId?: string; className?: string }) {
  const t = useT();
  const sheet = useRef<HTMLDivElement>(null);
  // Taken as the sheet renders, before anything in it takes the focus.
  const [opener] = useState(() => (typeof document !== 'undefined' ? document.activeElement : null) as HTMLElement | null);
  const [drag, setDrag] = useState<{ start: number; dy: number } | null>(null);
  useEffect(() => {
    const el = sheet.current;
    if (el && !el.contains(document.activeElement)) el.querySelector<HTMLElement>(FOCUSABLE)?.focus();
    return () => {
      // React's StrictMode unmounts once at mount without removing the element (development only): nothing to give back then.
      if (el?.isConnected) return;
      // Only when the focus went with the element: what took it in the same commit (a dialog's field, ED-013
      // section 9) keeps it.
      if (opener?.isConnected && focusIsNowhere()) opener.focus();
    };
  }, [opener]);
  const trap = (e: KeyboardEvent<HTMLDivElement>): void => {
    if (e.key === 'Escape') { e.stopPropagation(); onClose(); return; }
    if (e.key !== 'Tab' || !sheet.current) return;
    const items = [...sheet.current.querySelectorAll<HTMLElement>(FOCUSABLE)];
    if (items.length === 0) return;
    const first = items[0]!;
    const last = items[items.length - 1]!;
    if (e.shiftKey && document.activeElement === first) { e.preventDefault(); last.focus(); }
    else if (!e.shiftKey && document.activeElement === last) { e.preventDefault(); first.focus(); }
  };
  // A swipe starts anywhere on the handle or the title but a button: a captured pointer's click would go to the capturer,
  // and the close button would never be clicked.
  const down = (e: PointerEvent<HTMLDivElement>): void => {
    if ((e.target as Element).closest('button')) return;
    try { e.currentTarget.setPointerCapture(e.pointerId); } catch { /* a pointer the browser no longer has */ }
    setDrag({ start: e.clientY, dy: 0 });
  };
  const move = (e: PointerEvent<HTMLDivElement>): void => { if (drag) setDrag({ start: drag.start, dy: Math.max(0, e.clientY - drag.start) }); };
  const up = (): void => { if (drag && drag.dy > SWIPE_CLOSE) onClose(); setDrag(null); };
  const swipe = { onPointerDown: down, onPointerMove: move, onPointerUp: up, onPointerCancel: () => setDrag(null) };
  return (
    <div className="sheet-backdrop" onMouseDown={(e) => { if (e.target === e.currentTarget) onClose(); }}>
      <div ref={sheet} className={`sheet${className ? ` ${className}` : ''}`} role="dialog" aria-modal="true" aria-label={title} data-testid={testId} onKeyDown={trap}
        style={drag && drag.dy > 0 ? { transform: `translateY(${drag.dy}px)` } : undefined}>
        <div className="sheet-top" data-testid="sheet-top" {...swipe}>
          <div className="sheet-handle" aria-hidden="true" />
          <div className="sheet-head">
            <h2>{title}</h2>
            <button type="button" className="sheet-close" aria-label={t('sheet.close')} title={t('sheet.close')} data-testid="sheet-close" onClick={onClose}>×</button>
          </div>
        </div>
        <div className="sheet-body">{children}</div>
      </div>
    </div>
  );
}

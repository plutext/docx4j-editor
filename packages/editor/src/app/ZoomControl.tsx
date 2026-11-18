// SPDX-License-Identifier: GPL-3.0-or-later
//
// The status bar's zoom control (ED-005 section 12.110), as Word's: a minus, the percentage, a plus; the percentage
// returns the page to 100%.
import { useT } from './i18n.js';
import { ZOOM_MIN, ZOOM_MAX, stepZoom } from './zoom.js';

export function ZoomControl({ zoom, onZoom }: { zoom: number; onZoom: (zoom: number) => void }) {
  const t = useT();
  return (
    <span className="zoom" data-testid="zoom" role="group" aria-label={t('zoom.label')}>
      <button type="button" className="zoom-step" data-testid="zoom-out" aria-label={t('zoom.out')} title={t('zoom.out')}
        disabled={zoom <= ZOOM_MIN} onClick={() => onZoom(stepZoom(zoom, -1))}>−</button>
      <button type="button" className="zoom-level" data-testid="zoom-level" title={t('zoom.level.title', { zoom })}
        onClick={() => onZoom(1)}>{t('zoom.level', { zoom })}</button>
      <button type="button" className="zoom-step" data-testid="zoom-in" aria-label={t('zoom.in')} title={t('zoom.in')}
        disabled={zoom >= ZOOM_MAX} onClick={() => onZoom(stepZoom(zoom, 1))}>+</button>
    </span>
  );
}

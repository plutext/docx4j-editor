// SPDX-License-Identifier: GPL-3.0-or-later
//
// The perspectives this application ships (ED-003 section 4.9; ED-005 section 6.7): Edit with no side panel; Review with the tracked changes;
// Template with the Control, Controls, Data, OpenDoPE and Preview tabs over the basic toolbar with the control group, suggested by a document with
// OpenDoPE parts or a data binding; Developer with its seven tabs (the XPath tab since E4.a). The Template tabs and the
// console load with `import()` when first shown, so the Edit perspective's bundle stays what
// it is.
import { Suspense, lazy, type ReactNode } from 'react';
import { registerPerspective, type TabContext } from './perspectives.js';
import { XmlPane } from './panels/XmlPane.js';
import { DiagnosticsPanel } from './panels/DiagnosticsPanel.js';
import { PartsPanel } from './panels/PartsPanel.js';
import { OutlinePanel } from './panels/OutlinePanel.js';
import { FormatPanel } from './panels/FormatPanel.js';
import { StylesPanel } from './panels/StylesPanel.js';
import { t } from './i18n.js';

const Console = lazy(() => import('./panels/Console.js'));
const ControlPanel = lazy(() => import('./panels/template/ControlPanel.js'));
const ControlsTree = lazy(() => import('./panels/template/ControlsTree.js'));
const DataPanel = lazy(() => import('./panels/template/DataPanel.js'));
const OpenDopePanel = lazy(() => import('./panels/template/OpenDopePanel.js'));
const PreviewPanel = lazy(() => import('./panels/template/PreviewPanel.js'));
const XPathPanel = lazy(() => import('./panels/XPathPanel.js'));
const ReviewPanel = lazy(() => import('./panels/ReviewPanel.js'));

function loading(label: string): ReactNode {
  return <section className="panel" aria-label={label}><h2>{label}</h2><p>{t('tab.loading')}</p></section>;
}

/** A lazily loaded tab, its heading while it loads the tab's label (`tab.<id>`) in the current language. */
const deferred = (id: string, node: ReactNode): ReactNode => <Suspense fallback={loading(t(`tab.${id}`))}>{node}</Suspense>;

registerPerspective({
  id: 'edit', label: 'Edit', panel: { open: false, tab: '' }, toolbar: 'full', tabs: [],
});

// Review (ED-005 section 6.7): the tracked changes' panel over the basic toolbar. Suggested by the
// diagnostics banner's offer, not by `suggest`, so that a document is offered it once.
registerPerspective({
  id: 'review', label: 'Review', panel: { open: true, tab: 'changes' }, toolbar: 'basic',
  tabs: [
    { id: 'changes', label: 'Changes', render: (ctx: TabContext) => deferred('changes', <ReviewPanel editor={ctx.editor} session={ctx.session} onStatus={ctx.onStatus} onTrackChanges={ctx.onTrackChanges} />) },
  ],
});

registerPerspective({
  id: 'template', label: 'Template', panel: { open: true, tab: 'control' }, toolbar: 'basic', controls: true, controlTags: true, formattingMarks: true,
  tabs: [
    { id: 'control', label: 'Control', render: (ctx: TabContext) => deferred('control', <ControlPanel {...ctx} />) },
    { id: 'controls', label: 'Controls', render: (ctx: TabContext) => deferred('controls', <ControlsTree {...ctx} />) },
    { id: 'data', label: 'Data', render: (ctx: TabContext) => deferred('data', <DataPanel {...ctx} />) },
    { id: 'opendope', label: 'OpenDoPE', render: (ctx: TabContext) => deferred('opendope', <OpenDopePanel {...ctx} />) },
    { id: 'preview', label: 'Preview', render: (ctx: TabContext) => deferred('preview', <PreviewPanel {...ctx} />) },
  ],
  suggest: (session) => session.template.loaded && session.template.isTemplate,
});

registerPerspective({
  id: 'developer', label: 'Developer', panel: { open: true, tab: 'codes' }, toolbar: 'full', formattingMarks: true,
  tabs: [
    {
      id: 'codes', label: 'Reveal codes',
      render: (ctx: TabContext) => (
        <>
          <XmlPane editor={ctx.active} onOpenInConsole={ctx.onConsoleScript} />
          {deferred('console', <Console editor={ctx.editor} session={ctx.session} perspective={ctx.perspective} script={ctx.consoleScript} onSaveMacro={ctx.onSaveMacro} ask={{ onStatus: ctx.onStatus }} />)}
        </>
      ),
    },
    { id: 'format', label: 'Format', render: (ctx: TabContext) => <FormatPanel editor={ctx.active} session={ctx.session} /> },
    { id: 'styles', label: 'Styles', render: (ctx: TabContext) => <StylesPanel editor={ctx.active} session={ctx.session} /> },
    { id: 'parts', label: 'Parts', render: (ctx: TabContext) => <PartsPanel session={ctx.session} editor={ctx.editor} /> },
    { id: 'diagnostics', label: 'Diagnostics', render: (ctx: TabContext) => <DiagnosticsPanel editor={ctx.editor} session={ctx.session} /> },
    { id: 'outline', label: 'Outline', render: (ctx: TabContext) => <OutlinePanel editor={ctx.editor} session={ctx.session} /> },
    { id: 'xpath', label: 'XPath', render: (ctx: TabContext) => deferred('xpath', <XPathPanel editor={ctx.editor} session={ctx.session} onConsoleScript={ctx.onConsoleScript} />) },
  ],
});

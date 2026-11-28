// SPDX-License-Identifier: GPL-3.0-or-later
import { StrictMode } from 'react';
import { createRoot } from 'react-dom/client';
import { App } from './app/App.js';
import { languageReady } from './app/i18n.js';
import { stopFocusZoom } from './app/compact.js';
import { initHost, host } from './app/host.js';
import { setLinkOpener } from './view/plugins/links.js';
import './view/editor.css';
import './view/print.css';
import './app/app.css';
import './app/print.css';
import './app/compact.css';

// No zoom when a tap focuses the text on iOS (ED-013): before anything can take the focus.
stopFocusZoom();

// The host chosen before anything reads an option (ED-014 section 4.1): VS Code's fills the page's storage from its own copy.
initHost();
// Ctrl+click on a link opens it through the host: a new tab, or VS Code's `env.openExternal`, since a webview has no pop-ups.
setLinkOpener((url) => host().openExternal(url));

// The interface's language loaded first (ED-003 section 12.2): its messages are a chunk of their own. The person's choice in
// Options, else the host's (VS Code's display language, ED-014 P27), else the browser's.
await languageReady(host().language());

createRoot(document.getElementById('root')!).render(
  <StrictMode>
    <App />
  </StrictMode>,
);

// SPDX-License-Identifier: GPL-3.0-or-later
import { StrictMode } from 'react';
import { createRoot } from 'react-dom/client';
import { App } from './app/App.js';
import { languageReady } from './app/i18n.js';
import { stopFocusZoom } from './app/compact.js';
import './view/editor.css';
import './view/print.css';
import './app/app.css';
import './app/print.css';
import './app/compact.css';

// No zoom when a tap focuses the text on iOS (ED-013): before anything can take the focus.
stopFocusZoom();

// The interface's language loaded first (ED-003 section 12.2): its messages are a chunk of their own.
await languageReady();

createRoot(document.getElementById('root')!).render(
  <StrictMode>
    <App />
  </StrictMode>,
);

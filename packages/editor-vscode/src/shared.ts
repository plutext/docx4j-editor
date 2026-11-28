// SPDX-License-Identifier: GPL-3.0-or-later
//
// The files the extension takes from the other packages' source (ED-014 sections 4.3 and 4.5, P24): the application's
// protocol and page writer, and the companion's channel. None imports anything but the protocol, so esbuild bundles them as
// they are and both packages' typechecks cover them. Imported by a relative path, which needs no alias in either tool.
export * from '../../editor/src/app/host-protocol.js';
export { webviewPage, webviewPolicy, WebviewPageError, type WebviewPageOptions } from '../../editor/src/app/webview-page.js';
export type { FromCompanion, ToCompanion } from '../../editor-mcp/src/channel.mjs';
// The Ask box's one call to the Anthropic API, the same in both hosts (section 4.6, P45). It imports the SDK, which the
// extension's bundle then carries.
export { sendAnthropic, failureOf, chatPrompt, isCancelled, ANTHROPIC_ADDRESS, type SendOptions } from '../../editor/src/app/ask-send.js';

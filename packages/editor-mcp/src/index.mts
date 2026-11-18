// SPDX-License-Identifier: GPL-3.0-or-later
//
// `@docx4j/editor-mcp` (ED-005 section 8): the companion that lets an agent edit a document in a
// co-editing room beside a person, every edit a tracked change. Step F1: the process, the
// application, the relay and the probe served; a document shared into a room by a headless
// session. Step F3: the MCP tools of `docs/developer/agent.md`, over stdio and streamable HTTP.
export { createCompanion, AGENT_COLOR, type Companion, type CompanionOptions, type AgentRoom } from './server.mjs';
export { existingUnderRoots, writableUnderRoots, PathError } from './paths.mjs';
export { createMcpServer } from './tools.mjs';
export { ToolError, type ErrorCode } from './errors.mjs';
export type { ToolHost, ToolRoom } from './kit.mjs';

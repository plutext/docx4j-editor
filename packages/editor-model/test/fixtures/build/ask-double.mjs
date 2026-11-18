// SPDX-License-Identifier: GPL-3.0-or-later
//
// A double for the ask mode's provider (ED-003 section 7.5): the Messages API's shape, answered
// by the prompt's words, served under `/ask` by Vite's dev and preview servers beside the
// companion double so the browser tests prove the request preview, the three forms and the two
// insertions without a network or a key. "script" gets a ts block, "explain" prose, anything
// else a wml block holding the prompt's first line as a paragraph (styled Heading1 when the
// prompt says "heading").
//
//   POST <base>/v1/messages  -> { type: 'message', content: [{ type: 'text', text }], stop_reason, ... }

/** The answer text for a prompt. */
export function answerFor(prompt) {
  const first = (prompt.split('\n')[0] ?? '').trim();
  const lower = first.toLowerCase();
  if (lower.includes('refuse')) return null;
  if (lower.includes('script') || lower.includes('macro')) {
    return `Here is a script; read it and press Run.\n\`\`\`ts\n// ${first}\nfor (const p of selection.paragraphs) p.font.italic = true;\nreturn selection.paragraphs.length;\n\`\`\``;
  }
  if (lower.includes('explain')) return `The double explains: "${first}". A w:p is a paragraph, a w:r a run, a w:t the run's text.`;
  if (lower.includes('plain text')) return `\`\`\`text\nThe double's plain answer to: ${first}\nA second paragraph.\n\`\`\``;
  const escaped = first.replace(/&/g, '&amp;').replace(/</g, '&lt;');
  const style = lower.includes('heading') ? '<w:pPr><w:pStyle w:val="Heading1"/></w:pPr>' : '';
  return `A paragraph for the document:\n\`\`\`wml\n<w:p>${style}<w:r><w:t xml:space="preserve">The double answers: ${escaped}</w:t></w:r></w:p>\n\`\`\``;
}

/** The Messages API's response for a request body. */
export function messageFor(body) {
  const content = body?.messages?.[0]?.content;
  const prompt = typeof content === 'string' ? content : Array.isArray(content) ? content.map((c) => c.text ?? '').join('') : '';
  const text = answerFor(prompt);
  const model = body?.model ?? 'double';
  if (text === null) {
    return { id: 'msg_double', type: 'message', role: 'assistant', model, content: [], stop_reason: 'refusal', stop_details: { type: 'refusal', category: 'test', explanation: 'The double declines prompts that say refuse' }, stop_sequence: null, usage: { input_tokens: 1, output_tokens: 0 } };
  }
  return { id: 'msg_double', type: 'message', role: 'assistant', model, content: [{ type: 'text', text }], stop_reason: 'end_turn', stop_details: null, stop_sequence: null, usage: { input_tokens: 1, output_tokens: 1 } };
}

/** The route on a Node `(req, res, next)` handler under `base` (`/ask`), for Vite's dev and preview servers. */
export function askRoutes(base = '/ask') {
  return (req, res, next) => {
    const url = (req.url ?? '').split('?')[0];
    if (url !== `${base}/v1/messages`) { next(); return; }
    if (req.method !== 'POST') { res.statusCode = 405; res.end(); return; }
    const chunks = [];
    req.on('data', (c) => chunks.push(c));
    req.on('end', () => {
      let body;
      try { body = JSON.parse(Buffer.concat(chunks).toString('utf8')); } catch { res.statusCode = 400; res.end('{"error":"bad json"}'); return; }
      if (!req.headers['x-api-key'] && !req.headers['authorization']) { res.statusCode = 401; res.setHeader('Content-Type', 'application/json'); res.end(JSON.stringify({ type: 'error', error: { type: 'authentication_error', message: 'no key' } })); return; }
      res.setHeader('Content-Type', 'application/json');
      res.end(JSON.stringify(messageFor(body)));
    });
  };
}

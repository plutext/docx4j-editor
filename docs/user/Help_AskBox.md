# The Ask box

The Ask box is at the foot of the console, in the Developer perspective's Reveal codes tab. You
ask a question in plain language about the part of the document you have selected, and Claude
answers in a form the editor can act on. Nothing changes in the document until you press a
button.

## What Claude sees

Each question is sent with:

- the selected text, and its XML (WordprocessingML);
- the document's headings, as an outline;
- the document's style names and ids;
- the type declarations of the editor's scripting interface (the Office JS `Word.*` subset).

Claude does not see the rest of the document's text, its images, or how it is laid out on the
page. Before anything is sent, the box shows you the whole request.

## Four kinds of answer

### 1. New content for the document

Shown as a preview in the document's own styles; **Insert at selection** puts it after the
selected blocks, **Replace selection** puts it in their place. Examples:

- "Write a two-sentence summary of this section as a Heading 2 and a paragraph."
- "Turn this list into a three-column table with a header row."
- "Rewrite this paragraph more formally."
- "Add a numbered list of the steps described here."

Claude uses the document's style ids, so the result matches the document's look.

### 2. A script

Placed in the console editor above the box, **never run** until you press Run. Scripts suit
changes across the whole document that are tedious by hand:

- "Make every Heading 2 use the Heading 3 style instead."
- "Italicise every occurrence of 'et al.'"
- "Add a paragraph after each table with the text 'Source: …'."
- "Write a macro that bolds the first word of every paragraph."

Read the script, then press Run: the whole run is one undo step. A script you want again can be
kept with **Save as macro…**.

### 3. Plain text

Inserted at the caret (or replacing the selected text) when formatting does not matter.

### 4. An explanation

Shown in the box; it changes nothing:

- "Why is this paragraph indented?" (Claude sees the XML, direct formatting included.)
- "What does this content control's binding do?"
- "Explain this XML."

## What it is not good at

- **The whole document's text.** Only the selection and the outline are sent, so "summarise the
  whole document" works only if you select everything first.
- **Tracked changes, comments, images and page layout.** Claude cannot see images, and the
  layout is not in the XML it receives.
- **Very large selections.** A long selection makes a long request, a slow answer and, with an
  API key, a larger bill.

## Two ways to use it

### With an Anthropic API key

Enter your key in **Options** ("Ask: Anthropic API key"). It is kept in this browser and sent
only to the provider. Ask, check the request, and press **Send**; or tick "Send without showing
the request" to skip the check. The API is billed to your Anthropic account, separately from
any Claude plan.

### With a Claude plan, or no key at all (the free plan works)

1. Leave the key empty, type your question and press **Ask**.
2. Press **Copy prompt**. The whole request goes on the clipboard as one message.
3. Paste it into a chat at [claude.ai](https://claude.ai/new), or any other Claude chat.
4. Copy Claude's whole reply, including its fenced block (the part between the ``` lines).
5. Paste it into the box under the request and press **Use this answer**.

The pasted answer is handled exactly as one from the API: content is previewed and waits for
Insert or Replace, a script goes to the console unrun, an explanation is shown.

## Tips

- Select exactly what the question is about: the answer is only as good as what Claude sees.
- Name styles the way the document does ("Heading 2", "Quote"); the style list is in the request.
- For a change across the document, ask for a script rather than new content.
- Everything the box inserts is one undo step (Ctrl+Z).

# A picture node holding a data URL: ED-005 section 12.125's Word check

Built by `node packages/editor-model/test/fixtures/build/picture-dataurl-for-word.mjs` from the repository root. The XML is
written by hand, with no engine. Run in Word 365 through the portfolio's Word bridge (Jason, 2026-11-12: "try the word
bridge"), round `picture-dataurl`. Each file is opened and saved as docx and as PDF. Word's saves go beside the inputs as
`pdu1-base64-empty-word365.docx`, `.pdf` and so on. Every input is in compatibility mode 15.

## The question

The editor and the engine read a node holding `data:image/png;base64,...` as the base64 after the comma (core-ts
`customxml/pictures.mts`; section 12.125). Word writes bare base64 when it writes a picture control's image to its node. Does
Word read a data URL? What does it do with the control when it can't, and what does it write back to the node?

## The files

Each holds one paragraph of text, one picture control (`w:picture`, `w:id` 235915003, alias "Picture") mapped by
`w:dataBinding` to `/data/pic` in `customXml/item1.xml`, and "After the control.". The node's picture is a 120 by 60 orange PNG.

| File | The node holds | The control holds |
|---|---|---|
| `pdu1-base64-empty.docx` | bare base64 | the editor's placeholder text, `w:showingPlcHdr` |
| `pdu2-dataurl-empty.docx` | the same PNG as a data URL | the placeholder |
| `pdu3-dataurl-filled.docx` | the data URL | a 120 by 60 blue picture, as the editor's refresh writes one |
| `pdu4-base64-filled.docx` | bare base64 | the blue picture |

## What to read from Word's saves

- **pdu1 (the control):** Word fills a mapped picture control from its node on open (core-ts checks 37 and 38). Orange in the
  PDF.
- **pdu2:** orange means Word reads a data URL. The placeholder, a blank or a broken picture means it doesn't.
- **pdu4 (the control for pdu3):** orange means Word refills a filled control from its node on open; blue means it keeps the
  control's own picture.
- **pdu3:** with pdu4's answer, what an editor-made file shows in Word. The editor's real file holds the node's own picture in
  the control, not another colour.
- **In every save, `customXml/item1.xml`:** whether Word wrote the node back (base64 of which picture, or the data URL left as
  it was). The saved `document.xml` shows what the control holds.

## Run 1: Word 365 (2026-11-12, bridge 22479f7, Word 16.0.20326.20318)

Round `picture-dataurl`, with the control pair `pdu-mu-control-a` and `-b`. All six jobs opened silently and saved
completely, with no dialogs or repairs. Word's saves are beside the inputs as `*-word365.docx` and `.pdf`. The PDFs' colours
were read by rendering the pages; the pictures in the saves were compared by bytes with the inputs'.

| File | Word showed | The control in Word's save | The node in Word's save |
|---|---|---|---|
| `pdu1` (bare base64, placeholder) | the placeholder text | the text and `w:showingPlcHdr`, no drawing | as it was |
| `pdu2` (data URL, placeholder) | the placeholder text | the same | as it was |
| `pdu3` (data URL, blue picture) | blue | the blue picture's bytes | as it was |
| `pdu4` (bare base64, blue picture) | orange | the node's picture's bytes | as it was |

- **Word doesn't read a data URL.** Over one it keeps whatever the control holds, with no message (pdu2, pdu3).
- **Over bare base64 Word refills a control that holds a picture** (pdu4). It swaps the picture's bytes, as core-ts CR-005
  section 10 found.
- **A picture control holding placeholder text is never filled, even from bare base64** (pdu1). There is no drawing whose bytes
  Word could swap. The editor's picture control holds such text until its refresh writes a picture into it.
- **Word writes nothing back to the node** in any of the four.

`test/pictures.test.mts` pins these answers against the saves, and the editor's reading of each.

## Round 2: why Word left pdu1 (staged 2026-11-12, round `picture-placeholder`)

Jason, asked why Word filled pdu4 and not pdu1: "yes stage the round". pdu1's control held placeholder text with
`w:showingPlcHdr`, and pdu4's held a picture without it. Each of these holds bare base64 of the same orange picture.

| File | The control holds |
|---|---|
| `pdu5-base64-text-noflag.docx` | the placeholder text, with no flag |
| `pdu6-base64-placeholder-flag.docx` | Word's form as the ports write it (core-ts CR-005 section 10.1): the flag and a 100-point drawing of docx4j's `picture-placeholder.png` (SHA-1 `e9e4ee21`) |
| `pdu7-base64-empty-content.docx` | nothing, with no flag |

- **pdu5 filled:** the flag stopped pdu1.
- **pdu6 filled:** the flag doesn't block, and a picture to swap is what Word needs. That is the form the editor's new picture
  control would take.
- **pdu7:** what Word does with a control that holds nothing.

### Run 2 (2026-11-12, bridge 22479f7, Word 16.0.20326.20318)

The control pair `pdr2-mu-control-a` and `-b` opened silently. Word's saves are beside the inputs as `*-word365.docx` and
`.pdf`, with the repair dialog as `pdu5-word365-repair-dialog.png`.

| File | Opened | Word showed | The control in Word's save |
|---|---|---|---|
| `pdu5` (text, no flag) | **repaired**: "Word found unreadable content", once on each open, and the bridge answered Yes | the text | the text, kept |
| `pdu6` (Word's form, the ports' picture) | silent | orange | the node's picture in the placeholder's drawing, the flag cleared, 1270000 by 635000 EMU |
| `pdu7` (nothing) | silent | nothing | still empty |

- **Word fills a mapped picture control only by swapping the picture it holds,** with the flag (pdu6) or without it (pdu4).
  With no picture in the control there's nothing to swap, so nothing is filled, whether the control holds text (pdu1) or
  nothing (pdu7).
- **Text in a picture control without the flag is content Word repairs** (pdu5). With the flag Word keeps it silently
  (pdu1). This is where check 71's round 1 repair came from.
- **Over the placeholder Word keeps the drawing's width and fits its height to the picture's shape,** turning the 100-point
  square into 100 by 50 points for a 2:1 picture. core-ts's bind keeps the extent as it is, square.
- **So the editor's new picture control should hold Word's form,** the flag and the placeholder picture. Word then shows the
  node's picture on open.

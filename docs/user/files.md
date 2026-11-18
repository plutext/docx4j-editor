# Files on disk

How the editor opens and saves files, and what it does when a file it has open changes on disk.

## Opening

**Open…** uses the browser's file picker. In Chromium (Chrome, Edge) the picker gives the editor a
*handle* on the file, so that **Save** can write it in place and the editor can see the file change
on disk. Dropping a file onto the editor gives a handle in Chromium too. Firefox and Safari give no
handle: the file opens as usual, **Save** downloads a copy or asks where to save, and a change on
disk is not seen.

**Download** saves a copy through the browser's downloads, and **Save As** asks where to save. **Save**,
which writes the file in place through its handle, is an option, **Options → Save in place**, off by
default until it is verified with Word 2010; with it on, a **Save** button appears for a document
with a handle. Before writing, Save checks the file's modification time against the time it was
opened or last saved, and refuses with "changed on disk" when another program wrote the file
meanwhile, so that it never writes over that program's save unseen.

## Saving as Flat OPC XML

A document can be saved as Flat OPC: the whole package as one XML file, which Word opens as a "Word
XML Document" (`.xml`). In **Save As**, choose "Word XML document (Flat OPC)" and give the file a name
ending `.xml`. The command palette's **Save as Flat OPC XML…** suggests that name. Without a file picker
(Firefox, Safari), the same command downloads the document as `name.xml`.

The name decides the form: a name ending `.xml` is written as Flat OPC, and any other name as a `.docx`,
whatever type the picker shows. After that, **Save** and **Download** keep to the same form. The parts
are the ones a `.docx` save writes, and the save is checked by reading back the file it wrote.

The editor saves Flat OPC files but does not open them. A document saved as `.xml` is not watched for
changes on disk, since **Reload** would have to open it. **Save** still refuses with "changed on disk"
when another program wrote the file since.

## A file changed on disk

While a document with a handle is open, the editor reads the file's modification time every few
seconds, and at once when you come back to the tab. When the file was written, a banner says so, and
the banner's second sentence compares the text:

- "Comparing the two…" while it reads both versions, up to a second or two on a long document;
- then "On disk, 3 paragraphs are changed, 1 added and 0 removed", counting the paragraphs whose
  text differs between what the editor shows and what is on disk;
- or "The text of every paragraph is the same", which means the change is in formatting,
  properties or another part of the file (a header, a style, a setting);
- or that the disk version could not be read, for instance while another program is still
  writing it.

What the banner offers depends on whether you have edited the document here.

**No unsaved changes: Reload from disk, or Keep this version.** Reload opens the file as it is now,
keeping the perspective, the panel and your place (below). Keep closes the banner; it comes back if the
file is written again.

**Unsaved changes: Open the disk version in a new tab, Reload from disk discarding your changes, or Keep
this version.** If you do not want your changes, the second replaces them with the file as it is now, after
asking: they cannot be brought back afterwards, by Undo or by recovery. To keep both, the editor holds one
document per tab, so the disk version opens in a new browser tab, shown as "name (on disk).docx"
until it is saved in any way, and the document you were editing stays where it is with its changes.
The new tab opens at your place, found the same way as a Reload finds it, so that the two versions start
from the same line.
Compare the two by eye, copy what you need across, and save whichever you want to keep: **Save As**
or **Download** in the edited tab keeps your version without touching the file on disk, and in the
new tab the same two save the disk version under a name of your choosing (the file's own name is
suggested), or **Save** writes the file itself where the Save in place option is on. The first
**Save** in the new tab shows the browser's own question whether the site may edit the file, since
the tab got the file from the other tab rather than from a picker; allow it and the file is written,
refuse it and Save As remains. If the browser blocks the new tab, allow pop-ups for the site and
try again.

One thing the editor cannot do for you: if the other program still has the file open, it does not
see what the editor wrote. Word, for one, holds its own copy and does not watch a local file; its next
save writes over the editor's, and the editor's banner then appears again. Close the file in Word
before saving over it from the editor, or save under another name.

Nothing is merged: the editor shows you what differs and leaves the choice to you. Showing the disk
version's changes as tracked changes, to accept or reject, is planned (ED-011 phase 2) and waits on a
comparison in the engine.

## Your place after a reload

Reload keeps the view and the caret. If the caret is on screen, it comes back in the same text at the same
height on screen, even if paragraphs were added or removed above it. If you had scrolled away from the caret,
or never clicked in the document, the view stays on what you were reading and the caret goes back where it
was, out of sight. The view of the tracked changes (All markup, As accepted, Original) stays as it was too.

The place is found by the text, so what happens depends on what the other program did to it:

- **Its paragraph was edited:** the caret stays in the paragraph, with the words it was beside. A caret inside
  the words that changed goes to the start of the change.
- **Its paragraph was moved:** the caret follows it, as long as the text appears once and is more than a few words.
- **Its paragraph is gone:** the caret goes where the paragraph stood, at the start of the paragraph that now
  follows the gap, or at the end of the document if everything after that point went. The status line quotes
  the paragraph's opening words, since after a Reload nothing else shows what was removed.
- **Nothing in the two versions matches:** the caret and the view go to the start, and the status line says so.

A selection comes back as a caret. Only the document's own caret is kept: if you were in a note or a header,
it is the caret you last had in the document.

## Markdown files

A `.md` file opens as a new document, `notes.docx`, made from it. Where the browser gave a handle, the
editor watches the Markdown file as it would a document, and the same banner appears, naming `notes.md`.
It works one way only: nothing you do in the editor is written back to the `.md`, and the document has no
**Save** until it has a file of its own.

- **Reload from disk** converts the Markdown again and replaces the document with the result. As for any
  file, it is offered only when you have no unsaved changes, and it keeps your place.
- **Open the disk version in a new tab** converts the Markdown there. That tab usually cannot read the
  file itself, so it does not offer later changes, and its status line says so.
- The banner's comparison is of the text, so a change of formatting alone, such as a heading's level,
  reads as "the text of every paragraph is the same".
- **Save As** makes the document the `.docx` you saved: the editor watches that file from then on, and
  later changes to the `.md` are no longer offered. **Download** only saves a copy, so the `.md` is still
  watched; a Reload after it replaces what you see, and your edits stay in the copy you downloaded.

## In a room

A shared document (see [co-editing](co-editing.md)) is the room's, not the file's: everyone's
changes sit on the bytes the sharer put in, and a file written on disk by another program has none
of that history. So while a document is shared the file is not watched and no banner appears; the
sharer's **Save** still refuses when the file changed on disk, and Save As saves elsewhere.

## Recovery

Every open document is kept in the browser until it is saved or closed. If the browser is closed
with unsaved changes, the next visit offers to recover them. Close discards a document's unsaved
changes at once, and nothing is offered afterwards; to be asked first, turn on **Ask before closing a
document with unsaved changes** in Options, under All documents. A document still open in another tab
is not lost, so a new tab offers nothing; the offer comes once that tab has gone.

Two tabs, the edited document and the disk version, each keep their own, but the recovery banner
offers only the document opened most recently. If you work in two tabs and the earlier one crashes,
save the later one's work and do not count on the earlier one being offered.

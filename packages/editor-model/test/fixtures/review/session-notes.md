# Notes from the Word session (ED-005 section 9.2)

## Word 2010's id limit (2026-09-30, Jason, the M4600's Word 2010)

Each file in `ids/` opened in Word 2010. "OK" is opened with its footnote or comment; "refused" is
"cannot be opened because there are problems with the contents". A blank was not tried.

| Id | Footnote | Comment |
|---|---|---|
| 32,767 | OK | OK |
| 32,768 | refused | OK |
| 65,535 | | |
| 65,536 | | OK |
| 16,777,216 (2^24) | | OK |
| 536,870,912 (2^29) | refused | OK |

So Word 2010 takes a footnote id up to 32,767 (2^15 - 1, a signed 16-bit value) and refuses 32,768;
a comment id opens up to 2^29 at least. No second round is needed.

## The measurements

- `measure-word2010.docx`: Word 2010's save of `measure-base.docx`, sections 1 to 3 (2026-09-30).
- `measure-word15.docx`: Word 15's save of `measure-base.docx`, sections 1 to 5 (2026-09-30).

What both Words wrote (sections 1 to 3 agree between Word 2010 and Word 15):

1. **A comment on one word of Claude's pending insertion:** the anchor stays narrow.
   `<w:ins>The tenant pays rent <w:commentRangeStart/>monthly</w:ins><w:commentRangeEnd/>`, the
   reference run plain, then `<w:ins> in advance.</w:ins>`: the start marker inside the insertion,
   the insertion closed at the range's end, the reference outside it, the insertion split in two.
   (Word 2010's double-click took "monthly " with its space; Word 15's took "monthly".)
2. **Ctrl+L with tracking off** on a paragraph centred by direct formatting, its style left-aligned:
   the `w:jc` removed, the paragraph left with no `w:pPr`.
3. **New Comment at a caret, right to left:** 3a put the comment on the second word (עולם), 3b on the
   first (שלום), in both Words. **3c, an empty paragraph:** the reference run alone, no range markers.

Word 15 only, sections 4 and 5 (screenshots, and the markup section 4 left):

4. **A heading's paragraph mark deleted with tracking on:** the heading keeps Heading 2 with its mark
   deleted, and **the next paragraph takes Heading 2 as a `w:pPrChange`** by Author A (the old
   `w:pPr` empty). In No Markup the joined line is drawn in the heading's formatting throughout.
5. **Original over a formatting change:** "bold" shown not bold (the original formatting), with the
   change bar and the balloon "Formatted: Font: Bold" still drawn.

## The editor's file in Word 15 (2026-09-30)

- `review-for-word.docx`: the editor's file as built at 13:34 and opened by Jason (the same bytes
  in every copy he used).
- `review-for-word-word15.docx`: Word 15's save of it (Jason saved it over the original name);
  `review-for-word-word15-accepted.docx` and `-rejected.docx`: Accept All and Reject All, saved.

What the comparison found (the editor's reading and resolution against Word's, block by block: text,
paragraph style, alignment, indent, list, direction, runs' formatting, table rows, comments,
footnotes):

1. **Accept All and Reject All:** Word's documents are the editor's own Accept all and Reject all
   over the same file, with no difference; nothing pending in either; the five comments kept in
   both; the one footnote each way.
2. **The untouched save** keeps every change, comment (thread and resolved state) and footnote, and
   differs in two places:
   - a space inserted by "Author A" at the end of the second paragraph, tracked: a stray keystroke
     (the file's tracking is on, and Word's user name was still "Author A" from the measurements);
   - both hyperlinks, each holding a tracked change (Author B's link made over words, the agent's
     "lease"), rewritten by Word as HYPERLINK complex fields, as core-ts check 30 found Word keeps
     no `w:ins` inside a holder. The editor keeps the field as opaque content, and lists "lease"
     apart there, three changes, as Word's pane showed.

## The editor's file in Word 2010 (2026-09-30)

- `review-for-word-word2010.docx`: Word 2010's save of the editor's file (saved over its name);
  `review-for-word-word2010-accepted.docx` and `-rejected.docx`: Accept All and Reject All, saved.

The same comparison, the same result:

1. **Accept All and Reject All:** Word 2010's documents are the editor's own, with no difference;
   nothing pending; the one footnote each way; the five comments kept, in Word 2010's model (below).
2. **The untouched save** keeps every change and footnote, and differs in three places:
   - an extra space after "tools." in the second paragraph, not tracked: a stray keystroke, as in
     Word 15's save, in another place;
   - both hyperlinks rewritten as HYPERLINK fields, as Word 15 did: so both Words do it to a link
     holding a tracked change;
   - the comments in Word 2010's model, which has no threads: the reply a comment of its own and the
     thread's resolved state gone (ED-005 section 12.23 found the same).

**Both Words therefore resolve the editor's markup as the editor does.** The session's other
results are above; what is left is optional (current Word's saves).

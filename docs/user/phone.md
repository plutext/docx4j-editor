# Using the docx4j editor on a phone

On a phone the editor has a layout of its own, for simple edits and for review: typing and deleting text,
adding and deleting a table's rows, tracking changes, accepting and rejecting them, and commenting. It edits
the same document the full layout does, and saves the same file. Everything the full layout can do stays
reachable through **All commands** in the menu.

## When it applies

The phone layout is used on a touch screen with no mouse whose shorter side is at most 600 pixels (as the
browser counts them), in either orientation: a phone, not a tablet. **Options → Layout** in the menu
changes it: **Automatic (compact on a phone)**, **Compact** (also on a computer), or **Full** (the
computer's layout on a phone). The choice is kept in this browser.

## The top bar

One bar across the top: the menu (☰), the file's name (with • when it has unsaved changes), Undo and
Redo, the **Tracking** switch, and **Edit** | **Review**. While the keyboard is open the bar leaves out the
file's name and keeps the rest. **Tracking** and **Edit** leave the keyboard open; **Review** closes it to
show the changes. In landscape with the keyboard open there is no top bar at all, so that a few lines of
the document fit above the keyboard: type in portrait where you can.

The menu is a sheet from the bottom of the screen: **Open…**, **New**, **Save** and **Save As…** (in Chrome on
Android; see below), **Download**, **Send a copy…** (where the browser can), **Share** (with a document open, where
a relay is set up) and **Join…** for editing with others, **Find**, **All commands**, **Close**, and **Options**.
A sheet closes with its ✕, with a swipe down its top, or with Escape on a keyboard.

## The page

The page is drawn at the screen's width where the document's own text is wider: its lines break
differently from Word's, but the document is unchanged. In landscape the document's own width often fits,
and its lines then break as in Word. Wide tables scroll sideways inside themselves.

## Typing, and the bar above the keyboard

While you type, a bar sits just above the keyboard, with what fits where the caret is:

- **Hide the keyboard**, always, at the end: a keyboard over a down arrow, as Android shows it. It closes the keyboard
  and brings back the full top bar.
- In a table: **Row above**, **Row below** and **Delete row**.
- With text selected: **Comment**.
- On commented text: **Comments (n)**, the comments there.
- In Review: **Previous**, **Accept** and **Reject** with the caret in a change, and **Next**. Previous is greyed at
  the first change and Next at the last; they do not go round to the other end.

Its buttons leave the keyboard open. Samsung Keyboard (with Predictive text on or off) and Gboard have
been checked with tracking off and on, in Chrome on Android.

## Reviewing

Choose **Review** in the top bar: the document's tracked changes open in a sheet, each with **Accept** and
**Reject**, and with **Previous**, **Next**, **Accept all** and **Reject all** above them. Tap a change to
go to it: the sheet closes with the change selected, and with the keyboard open the bar above it offers
**Accept**, **Reject** and **Next**. **Accept all** and **Reject all** (and an author's all, under **By
author**) ask first, in a line under the button. Choose **Review** again to see the list again.

The **Tracking** switch records your own edits as tracked changes, under your name; the first time, it asks
the name. `docs/user/review.md` says more about tracked changes and comments in general.

## Comments

Select some text and tap **Comment** on the bar: a sheet shows the text you selected and a field for the
comment, and asks your name the first time. On commented text, **Comments (n)** opens its comments: each
with its author and date, **Reply**, **Resolve** (or **Reopen**), and **Delete** on your own. A comment's
text is not edited on a phone; the full layout does that.

## Files on a phone

- **Open…** uses the phone's own file picker (recent files, Downloads, Drive, iCloud Drive and the like).
- **Save** and **Save As…** are in the menu only in Chrome on Android, which has the file pickers they need.
  **Save** writes the file you opened, in place, when **Options → Save in place** is on; the first time,
  Chrome asks "Save changes to …?". **Save As…** asks where to save.
- **Download** saves a copy, in every browser. On an iPhone it is how you save: Safari puts the file in the
  Files app, under Downloads (unless Safari's settings name another folder).
- **Send a copy…** hands a copy of the saved document to the phone's share sheet, for mail, a chat or
  AirDrop, where the browser allows it, as Safari on an iPhone does. Chrome, Edge and
  Samsung Internet do not let a page send a Word document, so it is not in their menu: download the
  document, then share it from Chrome's download message or from the Files app.
- **Send a copy… needs the editor served over HTTPS.** Browsers allow sharing only from a secure page, so an
  editor opened at an `http://` address, such as a server on your own network, never offers it.
  `docs/user/files.md` says more.

## Not yet checked

- **iPhone.** Typing in portrait and landscape has been tried in Safari on an iPhone 13 Pro. The sheets, Review and
  a smaller iPhone have not, nor has an iPad. Safari's keyboard covers the page rather than shrinking it, so with the
  keyboard open the bars fade while you scroll and come back when the page settles. In landscape the editor draws no
  bars while you type (Safari's ✓ closes the keyboard) and the keyboard has no suggestions line, autocorrect or
  spelling underlines. Turning the phone while typing closes the keyboard: tap the text again to carry on. A tap puts
  the caret at the start or end of a word, as everywhere on an iPhone; touch and hold to place it inside one. Keep
  Safari's page zoom (the **aA** button) at 100%: above it the screen's right edge is cut off.
- **Other Android browsers** than Chrome, and other keyboards than Samsung Keyboard and Gboard.

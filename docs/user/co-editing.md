# Co-editing

Two or more people can edit one document at the same time, each in their own browser, and see
each other's changes as they are made. The document goes through a **relay**, a small server
that holds the shared copy and passes changes between the people in it.

## Sharing a document

1. Set the relay's address under **Options → Relay address**. A deployment may have set one
   already, and then there is nothing to do. With no address there is no **Share** button.
2. Open the document and press **Share**. The editor makes a new room on the relay, puts the
   document in it, and shows a link with a **Copy link** button.
3. Send the link to the people you want to work with.

The status in the application bar says where the connection is: **Shared**, **Connecting…**,
**Offline: editing locally**, **Refused by the relay**, **Room not found**, or **Room full**. Its
tooltip names the relay. **Link** shows the link again. If the relay itself stops while your
network stays up, the status can say **Shared** for up to 30 seconds before it notices; what you
type meanwhile is kept and sent when the connection is back.

## Seeing each other

**Options → Your name** is what the others see, with a colour made from it; change it at any
time and the room sees the new name at once. Without one you are "Guest" and a number.

- **Carets.** Each person's caret is a bar in their colour, their name above it when you point at
  it; a selection is a tint. A caret in a footnote, an endnote, a header or a footer shows there
  and not in the body.
- **The people in the room.** Their initials, in their colours, are in the application bar;
  pointing at one says who it is and where they are ("in the body", "in footnote 3", or "not in
  the text" for someone who has not put a caret anywhere yet). A click takes you to their caret,
  or says where they are when that note or header is not on screen. A screen reader hears who
  joins and leaves.
- **Leaving.** Someone who closes their page disappears at once. Someone whose connection drops
  stays listed, caret and all, for up to 30 seconds.

## Joining

Open the link, or press **Join…** and paste it. Before anything connects, the editor asks
whether to join through the relay the link names. A link can name any relay, so opening a
stranger's link would otherwise connect your browser to their server. Say **Join**, and the
document opens once the relay has sent it. A link to a room the relay does not hold (mistyped, or
deleted after a long time unused) is said as such; joining never makes a room, only sharing does.

## What everyone should know

- **Anyone with the link can edit.** There are no roles. A link cannot be withdrawn: to exclude
  someone, share the document again, which makes a new room with a new link.
- **The relay can read what is shared.** It holds the document and every change to it. Share
  through a relay run by someone you trust with the document: your own machine, your
  organisation's server. The Share dialog names the relay.
- **Your name and where your caret is go to the relay and to everyone in the room**, and nowhere
  else: nothing about who was present is written into the document.
- **Undo is your own.** Ctrl+Z takes back your changes, never a colleague's.
- **Offline, you keep editing.** Your changes are kept in the browser and sent when the connection
  is back; the others' arrive then too. If the browser is closed first, reopening the editor
  offers to recover the session and rejoins the room.
- **Everyone's save is the same file.** Download or Save As on any side writes everyone's changes,
  and two people saving the same state get the same bytes. The document counts as changed when a
  colleague types, since your next save will include what they wrote.
- **Scripts and macros run only where you run them.** A console script, a macro or an answer
  from the Ask box runs in your own page, and what it changes reaches the others as ordinary
  edits. Nobody's code runs in anyone else's page. A macro saved into the document by a colleague
  is, for you, a macro that came with the document: it is shown to you before its first run, as
  any document macro is.
- **The file on disk is not watched while the document is shared.** A file another program writes
  has none of the room's history, so the "changed on disk" banner of [files.md](files.md) does not
  appear in a room; the sharer's Save still refuses when the file changed, and Save As saves elsewhere.
- **Re-paginate moves only the page markers** while the document is shared. A new custom XML data
  part, a template's first OpenDoPE entry and the first macro saved into the document can all be
  made while shared, and reach everyone in the room.

## Running a relay

```
npx docx4j-editor-relay --port 4175 --data ./rooms
```

The relay listens on this machine only (`127.0.0.1`) unless `--host` says otherwise, accepts pages
from `localhost` unless `--origin` names the site that serves the editor (reachable from other
machines, it belongs behind a TLS proxy, given to the editor as a `wss:` address), keeps each room in a
file under `--data`, refuses a room past `--max-room` bytes (200 MB by default), and deletes a room
nobody has changed for `--idle-days` days (30 by default). A room is made only by Share: its
token becomes the room's, and only a hash of it is kept; a join to a room the relay does not hold
is refused, never made.
`packages/editor-relay/README.md` has the details, and how to embed the relay in another server.

The design is ED-004's section 4 (the co-editing relay); `docs/developer/architecture.md` describes the layers a
room replicates.

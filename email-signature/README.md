# Outlook signature — Shafick Hassan, Vitacare Pharmacy Rondebosch

The Vitacare Rondebosch signature, rebuilt as an installable Outlook signature. "Responsible
Pharmacist / Shafick Hassan (M.Pharm)" is rendered **into the same image as the Vitacare logo**,
so the name and the mark travel together as one graphic.

## What is in here

| File | What it is |
|---|---|
| `Shafick Hassan.htm` | The signature — HTML, table-based, built for Outlook's rendering engine |
| `Shafick Hassan.txt` | Plain-text fallback, used when a mail is sent as plain text |
| `Shafick Hassan_files/vitacare-rondebosch-logo.png` | Source logo. **Placeholder** — replace it, see below |
| `Shafick Hassan_files/vitacare-rondebosch-lockup.png` | The name + logo banner the signature displays. Generated, never edited by hand |
| `tools/build_lockup.py` | Builds the lockup from the logo |
| `signature-preview.png` | What the finished signature looks like |

## Step 1 — supply the logo

The logo file in `Shafick Hassan_files/` is a dashed placeholder box, not the Vitacare logo.
The only copy available while this was built was a photograph of a screen, with the mouse
pointer sitting across the figure icon at the exact point where the arm meets the body.
Rebuilding what the pointer covers would have meant redrawing the Vitacare mark, so it was
left out rather than guessed at.

The clean original is already in your mailbox:

1. Open any email carrying the signature — the thread from Ali Chicktay, or anything in your
   **Sent Items**.
2. Right-click the Vitacare logo in the message body.
3. Choose **Save as Picture**.
4. Save it as `vitacare-rondebosch-logo.png`, replacing the placeholder.

Around 640 px wide is ideal: crisp on a high-resolution screen, displayed at 320 px.

## Step 2 — rebuild the lockup

The signature does not display the logo directly. It displays the lockup — the banner holding
the name, the title and the logo together. Once the logo is replaced, the lockup has to be
rebuilt from it:

```bash
python3 tools/build_lockup.py "Shafick Hassan_files/vitacare-rondebosch-logo.png" \
    --layout beside \
    --out "Shafick Hassan_files/vitacare-rondebosch-lockup.png"
```

This needs Python with Pillow, plus the Carlito and Caladea fonts. If that is not to hand,
send the logo file over and the finished lockup comes back built.

Two arrangements are available:

| `--layout` | Arrangement | Notes |
|---|---|---|
| `beside` | Name on the left, logo on the right, one full-width banner | Matches the original signature. This is the default |
| `under` | Logo on top, name right-aligned beneath it | Narrower; sits in a column next to the contact details |

The lockup renders at double size and displays at half, so it stays sharp on high-resolution
screens. It is drawn on a transparent background and sits on whatever the mail client uses.

## Step 3 — install it

### Outlook for Windows (classic desktop)

1. Close Outlook.
2. Press `Win` + `R`, paste `%APPDATA%\Microsoft\Signatures` and press Enter.
3. Copy `Shafick Hassan.htm`, `Shafick Hassan.txt` and the whole `Shafick Hassan_files` folder
   into it.
4. Reopen Outlook: **File > Options > Mail > Signatures**.
5. "Shafick Hassan" appears in the list. Set it under **New messages** and under
   **Replies/forwards**, then **OK**.

### New Outlook for Windows, Outlook on the web, or Outlook for Mac

These store signatures in the cloud rather than in a folder on the machine, so the file-based
method does not apply ([Microsoft Q&A](https://learn.microsoft.com/answers/a/12180436)):

1. Open `Shafick Hassan.htm` in Edge or Chrome (double-click it).
2. Click into the **rendered page** and press `Ctrl` + `A` then `Ctrl` + `C` (`Cmd` on a Mac).
   Copy what you can see, not the HTML source — pasting source is the usual reason a paste
   silently does nothing ([Microsoft Q&A](https://learn.microsoft.com/answers/a/12646271)).
3. Go to **Settings > Mail > Compose and reply** (Mac: **Outlook > Settings > Signatures**).
4. Create a signature named "Shafick Hassan", click into the editing box, press `Ctrl` + `V`.
5. Choose it for new messages and for replies, then **Save**.

The lockup travels with the paste, so no separate upload is needed.

### Which method applies to you

| | Classic Outlook for Windows | New Outlook / OWA / Mac |
|---|---|---|
| Where signatures live | `%APPDATA%\Microsoft\Signatures` on the PC | In the cloud, on your Microsoft account |
| Install method | Copy the files in | Copy and paste the rendered signature |
| Files used | `.htm` + `.txt` + image folder | None — Outlook stores its own copy |
| Syncs to your other devices | No | Yes |
| Outlook must be closed first | Yes | No |

## Step 4 — check it

Send yourself a test mail and confirm:

- the banner renders rather than showing a red cross or an empty frame
- the name and title read cleanly at normal zoom
- nothing wraps onto a second line
- the orange strap sits flush under the contact block
- the phone numbers and email addresses are clickable

## Notes

### The name is an image now

That is what makes the name part of the logo, and it carries one consequence worth knowing:
a recipient whose mail client blocks images will not see the banner. Two things cover that:

- The image's `alt` text reads "Responsible Pharmacist - Shafick Hassan (M.Pharm) - Vitacare
  Pharmacy Rondebosch", which is what Outlook shows in place of a blocked image.
- `Shafick Hassan.txt` carries the full details for plain-text sends.

The contact details were deliberately left as live text rather than folded into the image, so
the phone numbers and email addresses stay clickable and searchable.

### Fonts

The signature text uses Calibri, with Cambria for the title, in line with the Vitacare house
style. The banner has its lettering baked in using Carlito and Caladea — open fonts built to
the same metrics as Calibri and Cambria, and visually indistinguishable at this size. They are
used because the text has to be drawn into the image at build time rather than picked up from
the reader's machine.

### Markup

The layout is a table with inline styles on purpose. Outlook for Windows renders mail through
Word, which ignores flexbox, grid and external stylesheets. Rewriting this with `<div>`s will
collapse it.

### Colours

Orange `#C0421A`, name `#1A1A1A`, contact details `#3A3A3A`, services line `#6E6E6E`. The
rounded ends on the orange strap come from `border-radius`; Outlook for Windows squares them
off, which is harmless, and every other client shows the pill shape.

### Keeping it current

If the name, title, phone number or a service line changes, update `Shafick Hassan.htm` and
`Shafick Hassan.txt` together. A change to the name or title also means editing the strings in
`tools/build_lockup.py` and rebuilding the lockup — the banner is an image, so it will not
update on its own.

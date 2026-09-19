# Outlook signature — Shafick Hassan, Vitacare Pharmacy Rondebosch

A rebuild of the Vitacare Rondebosch signature as a proper Outlook signature file, matching
the layout, wording and colours of the version in the existing email thread.

## What is in here

| File | What it is |
|---|---|
| `Shafick Hassan.htm` | The signature itself — HTML, table-based, built for Outlook's rendering engine |
| `Shafick Hassan.txt` | Plain-text fallback, used when a mail is sent as plain text |
| `Shafick Hassan_files/vitacare-rondebosch-logo.png` | **Placeholder.** Replace with the real logo before use — see below |
| `signature-preview.png` | What the finished signature looks like |

## Step 1 — supply the logo (required)

The logo file in `Shafick Hassan_files/` is a dashed placeholder box, not the Vitacare logo.
The only copy available while this was built was a photograph of a screen, which carries the
mouse pointer sitting across the figure icon and the texture of the monitor. Rebuilding the
part of the logo hidden under the pointer would have meant redrawing the mark, so it was left
out rather than guessed at.

The clean original is already sitting in your mailbox. To pull it out:

1. Open any email that carries the signature — the thread from Ali Chicktay works, or anything
   in your own **Sent Items**.
2. Right-click the Vitacare logo in the message body.
3. Choose **Save as Picture**.
4. Save it as `vitacare-rondebosch-logo.png` into the `Shafick Hassan_files` folder, replacing
   the placeholder.

Keep the filename exactly as it is and the signature will pick it up with no further edits.
Aim for roughly 640 px wide — that gives a crisp logo on a high-resolution screen while the
signature displays it at 320 px.

## Step 2 — install it

### Outlook for Windows (classic desktop)

1. Close Outlook.
2. Press `Win` + `R`, paste `%APPDATA%\Microsoft\Signatures` and press Enter.
3. Copy `Shafick Hassan.htm`, `Shafick Hassan.txt` and the whole `Shafick Hassan_files` folder
   into that folder.
4. Reopen Outlook and go to **File > Options > Mail > Signatures**.
5. "Shafick Hassan" now appears in the list. Set it under **New messages** and under
   **Replies/forwards**, then click **OK**.

### New Outlook for Windows, Outlook on the web, or Outlook for Mac

These do not read the Signatures folder, so the signature goes in by copy and paste:

These store signatures in the cloud rather than in a folder on the machine, so the file-based
method above does not apply to them ([Microsoft
Q&A](https://learn.microsoft.com/answers/a/12180436)). The signature goes in by copy and paste:

1. Open `Shafick Hassan.htm` in Edge or Chrome (double-click it).
2. Click into the **rendered page** and press `Ctrl` + `A` then `Ctrl` + `C` (`Cmd` on a Mac).
   Copy what you can see, not the HTML source — pasting source is the usual reason a paste
   silently does nothing ([Microsoft
   Q&A](https://learn.microsoft.com/answers/a/12646271)).
3. In Outlook go to **Settings > Mail > Compose and reply** (Mac: **Outlook > Settings >
   Signatures**).
4. Create a signature named "Shafick Hassan", click into the editing box and press `Ctrl` + `V`.
5. Choose it for new messages and for replies, then **Save**.

The logo travels with the paste, so no separate upload is needed.

### Which method applies to you

| | Classic Outlook for Windows | New Outlook / OWA / Mac |
|---|---|---|
| Where signatures live | `%APPDATA%\Microsoft\Signatures` on the PC | In the cloud, on your Microsoft account |
| Install method | Copy the files in | Copy and paste the rendered signature |
| Files used | `.htm` + `.txt` + image folder | None — Outlook stores its own copy |
| Syncs to your other devices | No | Yes |
| Outlook must be closed first | Yes | No |

## Step 3 — check it

Send yourself a test mail and confirm:

- the logo renders rather than showing a red cross or an empty frame
- nothing wraps onto a second line
- the orange strap sits flush under the details block
- the phone numbers and email addresses are clickable

## Notes

- Fonts are Calibri for the details and Cambria for "Responsible Pharmacist", in line with the
  Vitacare house style. Carlito and Caladea also appear in the font list purely so the layout
  previews correctly on Linux; Windows never reaches them.
- The layout is a table with inline styles on purpose. Outlook for Windows renders mail through
  Word, which ignores flexbox, grid and external stylesheets. Rewriting this with `<div>`s will
  collapse it.
- Colours: orange `#C0421A`, name `#1A1A1A`, details `#3A3A3A`, services line `#6E6E6E`.
- Rounded ends on the orange strap come from `border-radius`. Outlook for Windows squares them
  off, which is harmless; every other client shows the pill shape.
- If a phone number, address or service line ever changes, edit `Shafick Hassan.htm` and
  `Shafick Hassan.txt` together so the HTML and plain-text versions stay in step.

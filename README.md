# Shift Calendar

Turns the monthly shift spreadsheet (the sheets named like `Oct 26 A M`) into a calendar, and downloads anyone's schedule as an `.ics` file for Apple, Google or Outlook Calendar.

It is a static site: no server, no build step, no accounts.

## What it does

- Reads every sheet whose name matches `A M` / `A and M` (change the pattern in Settings).
- Month and week views. Pick whose schedule to show (defaults to **Bernard C**); **Show coworkers** adds anyone else, and the week view lays everyone out as a rota board.
- **Download .ics**: choose from today onward, the month on screen, or everything; optional reminder; location defaults to `2263 Fuller Ct, Ann Arbor, MI`. Times are written in `America/Detroit` with daylight-saving rules, so overnight shifts and the November clock change come out right.
- Remembers the last uploaded file (IndexedDB) and your choices (localStorage), both in the browser only. The spreadsheet is never uploaded anywhere.

## Reading the spreadsheet

| Code | Meaning | Start - end |
|------|---------|-------------|
| `A` | from the sheet legend | 3:30 PM - 12 AM |
| `2` | from the sheet legend | 2 PM - 10:30 PM |
| `E` | from the sheet legend | 6 AM - 2 PM |
| `M` | legend says "8-8", read as overnight | 8 PM - 8 AM next day; **10 PM - 8 AM when it starts on a Saturday or Sunday** |
| `D` | legend says "DAYS", no hours: **assumed** | 8 AM - 4:30 PM |
| `S` | special shift, no time listed | all-day event |
| `L` / `X` | time off requested / called in sick | shown on screen; exported only if you tick the box |
| `0`, `-`, `O`, blank | off | nothing |
| `3:30-8` | a time written in the cell | read as 3:30 PM - 8 PM |
| `4` | a lone number is a start time | starts at 4 PM and ends at that person's usual end time (a late start for an `A` person) |

The legend doesn't give hours for `D`, so check that one in **Settings**. Codes the app can't read are listed under "things to check" and left out of the `.ics`.

A person's "usual end" comes from the shift code they work most often. Settings also has a box for the same person written two ways; it ships with `Bernard = Bernard K`, because June's sheet lists him without the initial.

The month and year come from the sheet name (`Sept 26 - A M` is September 2026), not from cell A1, which is sometimes stray text.

## Host it on GitHub Pages

1. Create a repository and upload everything in this folder (drag the files into the GitHub web page, or `git push`).
2. Settings > Pages > Build and deployment > Deploy from a branch > `main` / `(root)` > Save.
3. After a minute the site is at `https://<you>.github.io/<repo>/`.

Don't commit the spreadsheet. `.gitignore` already excludes `*.xlsx`, and the repository would be public. You upload the file in the browser instead.

## Host it on Vercel

Import the repository (or run `vercel` in this folder). Leave the framework as **Other** and leave the build command and output directory empty.

## Notes for tinkering

- `core.js` holds the parsing and `.ics` logic with no browser dependencies. `app.js` is the UI. `style.css` is the look.
- `vendor/xlsx.full.min.js` is SheetJS 0.18.5, included so the site works without a CDN.
- Tests: `npm install && npm test`.
- Saved data is per browser and per site address. A different browser, or a new site address, asks for the file once.

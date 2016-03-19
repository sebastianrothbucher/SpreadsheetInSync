# SpreadsheetInSync
Real-time collaboration in/between two spreadsheet applications plus web interface with CouchDB (or Cloudant)

## What can you do with it?
(to come)

## Installing and first steps

### OpenOffice.org Calc (OOo)
(to come)

### Excel 
(to come)

## Web view
(to come)

### Installing directly

### Installing via install DB

### Using a Login database
(to come)

## Known limitations / things 2 keep in mind
(although you're using it at your own risk anyway)
- requires a recheck of the sheet after inserting / deleting rows or colums
- generally, a regular recheck of the sheet makes sense
- currenly uses the WinHttp COM libraries - hence runs on Windows systems only
- continuously issues requests to the CouchDB to stay updated (Excel version: with longpoll and OOo version: periodically) - these might be charged depending on your plan
- likewise: initial share (and recheck) might produce a significant amount of requests to the CouchDB (at least one per cell, possibly several)
- at least the OOo version is not reliable with more than one sheet at a time
- binaries have been scanned with an up-to-date Antivirus software - please do recheck anyway before running
- when checking SSL on/off, the port won't adjust automatically (unless you have other info: it's likely 80 for HTTP and 443 for HTTPS)
- there is no automated regression test for the Excel version yet
- build is currently manual
- Times displayed in the history might not be timezone-adjusted correctly yet

### for Web view specifically
- changed formula results might only show up after recheck of the sheet
- Web view edits without updating an actual Spreadsheet file will never result in formula changes
- any edit in the Web view are texts, so formulas might pick up your changes only after converting to numbers
- Web view just displays the first sheet available - yet without the ability to change that
- Installing the Web view via menu is only possible if you create an installation DB on your own system. Otherwise, you can install via install.sh directly
- columns that have lots of long text might grow disproportionally large
- navigating a filtered sheet in cursor keys might still result in the cursor disappearing (as filtered rows are not skipped)
- filtering does not span several cells of a row
- filtering can not contain comparison operators or be limited to certain cols
- there might not always be feedback on success of a change
- the Excel version does not color changes (yet)
- when using a separate DB to manage login to a sheet (useful e.g. for mobile), this URL is not shown via menu

## License
all code is licensed under: Apache License 2.0 (see LICENSE)


# SpreadsheetInSync
Real-time collaboration in/between two spreadsheet applications (Excel and OpenOffice.org Calc) plus web interface with CouchDB (or Cloudant). You keep working in the application you know, and you keep your files - but you have team sync in real time now!

## What can you do with it?
The plugin (for Excel and OpenOffice.org Calc) allow parallel work on a spreadsheet with copies simultanously open on several computers. Whatever structure and setup you need to get the job done, you're free to do it!

With the plugin, everyone will see every change instantly and is always up to date. As you keep using files, you always have all the info with you - also to work offline and sync back when you have network connection. 

It's up to you whether to use the Cloud or your own network.

Here's some of what's in 4u: 
- you're free to create the structure you need, the sheet is yours - literally
- real offline capability: work offline and sync back later
- real real-time: changes sync instantly 
- works in both local network and the Cloud
- audit trail of each edit
- just go back to file w/out plugin at any time (no risk of locking yourself in)
- smart extras like iCal, Chat, Web view

You can use the plugin for things like:
- agile board / issue tracking
- organizing an event or conference
- managing timesheets, contracts and budgets
- info you need on the road
- hire shortlist / personal development plans
- communication plans
- brainstorming and shortlisting
- backlog / features / priorization for development
- purchasing and logistics
- leads you want to follow up on
- shift plans

and, off course: 
- your personal TODO list

(and certainly a lot more)

## How it works
Basically, there is a plugin extending the OpenOffice.org Calc or Excel that you have. The plugin listens for changes to cells and sends them up to a hub database. Either at once or when you re-connect and recheck/resync the sheet.

So here's the steps: 

<img src="files/howitworks1.png" /> 

Alice does some work in a spreadsheet which the plugin running. All changes are pushed up to a hub database (CouchDB).

<img src="files/howitworks2.png" />

Bob and Chris have their copy of the file open with the plugin running...

<img src="files/howitworks3.png" />

... which makes them get Alice's changes instantly - so everyone is up to date all the time.

## Installation and first steps

### OpenOffice.org Calc (OOo)
First of all, you need a [CouchDB](http://couchdb.apache.org) installation on your machine, your network or the Internet. You can use [Cloudant](http://cloudant.com) but be sure you understand their mechanics of charging - there's limits to what you get for free that also might change. Create a new and empty database in either case to hold your data.

Then, download 'SpreadsheetInSync.oxt' from [Releases](https://github.com/sebastianrothbucher/SpreadsheetInSync/releases) and double-click to install into the Extension Manager. For every spreadsheet file you create or open after that, you have a 'SpreadsheetInSync' menu below Tools > Add-ons (or Extras > Add-ons for e.g. the German localization). Here's how it looks like (w/ German OOo): 
<img src="files/screen_ooo.png" />

Then, choose SpreadsheetInSync > Start. As it's the first time, it will prompt for the database details. Give server name (for Cloudant: 'yourname'.cloudant.com), port (standard for HTTP w/out SSL is 80, standard for HTTPS w/ SSL - the recommended way - is 443) and the name of the database you just created. 

Then (for the first time and any time you come back), it will prompt for username and passwort to CouchDB (for Cloudant it's per se the same credentials you log on to the website). When you run CouchDB in AdminParty w/out further protection (per se not recommended!), leave the password empty.

As soon as you have started, it will watch the database for any changes made by someone else. They will be inserted into the sheet and (for OOo) be marked by a yellow background color that disappears after a few seconds. Likewise, anything you type or change will be replicated up to the database. 

Rechecking the sheet is useful to make sure really all changes are uploaded.

Before closing the file, make sure to stop the replication.

### Excel 
First of all, you need a [CouchDB](http://couchdb.apache.org) installation on your machine, your network or the Internet. You can use [Cloudant](http://cloudant.com) but be sure you understand their mechanics of charging - there's limits to what you get for free that also might change. Create a new and empty database in either case to hold your data.

Then, download 'SpreadsheetInSync.xlam' from [Releases](https://github.com/sebastianrothbucher/SpreadsheetInSync/releases) and copy it into C:\Users\'yourname'\AppData\Roaming\Microsoft\AddIns. Open the Options dialog, choose the Add-ins tab and (on the bottom) 'Go to' manage Excel Add-Ins. A dialog opens where you can check 'Spreadsheetinsync'. You now have a new tab for SpreadsheetInSync.

Then, choose SpreadsheetInSync > Start. As it's the first time, it will prompt for the database details. Give server name (for Cloudant: 'yourname'.cloudant.com), port (standard for HTTP w/out SSL is 80, standard for HTTPS w/ SSL - the recommended way - is 443) and the name of the database you just created. 

Then (for the first time and any time you come back), it will prompt for username and passwort to CouchDB (for Cloudant it's per se the same credentials you log on to the website). When you run CouchDB in AdminParty w/out further protection (per se not recommended!), leave the password empty.

As soon as you have started, it will watch the database for any changes made by someone else. They will be inserted into the sheet. Likewise, anything you type or change will be replicated up to the database. 

Rechecking the sheet is useful to make sure really all changes are uploaded. 

Before closing the file, make sure to stop the replication.

## First replication

### Empty file, empty database
You'll be prompted for database details - and that's it ;-)

### Empty file, database with content
You'll be prompted for database details - and then receive all the content of the database into your file.

### File with content, empty database
You'll be prompted for database details - and you can then opt to push all changes "made offline" to the database. If you opt yes, you'll have both in sync. Depending on the size of the sheet, you might generate some traffic and it might take up some time.

## Working offline
Essentially, this is the generalization of the cases above: As soon as you start, you'll receive all changes pushed to the hub database while you were away. You'll then be prompted to upload all your changes since disconnecting. As you had the file in the meantime, everything is as usual.

## When you decide to throw out SpreadsheetInSync...
... which hopefully will never happen: you still have your files (the .xls/.xlsx/.ods files), so you can keep on working with them just like you did before, no harm done, no migration work necessary. 

## Web view
The Web view is a very simple spreadsheet you can open in the Browser: it runs on the hub database shows the contents you've uploaded there so far. You can also edit cells or filter for rows containing some values. 

### Installing directly
Once you have a hub database set up, you can install by cloning this repo, pointing your bash to the 'webview' subdirectory and running
```
./prereq.sh && ./install.sh <path2hubdb>
```

For instance (given server is couch.local and DB is mysheet)
```
./prereq.sh && ./install.sh 'http://couch.local:80/mysheet/'
```

This will download all libraries needed (happens only once) and then upload all that's necessary for the Web view via the following URL: 
```
http://couch.local:80/mysheet/_design/showfkt/_show/htmlout/webview
```
(There's also a menu item in OpenOffice.org Calc that gives you the Web view URL copy&paste-ready)

### Installing via install DB (OOo only)
To give users a somewhat more convenient way to install the Web view: create a database named 'webview_install' and install the Web view there as described above (don't actually use this DB as a hub DB). Then, choose SpreadsheetInSync > Webview > Install Webview from the menu

### Using a Login database
(to come)

## Advanced features

### History (OOo only)
(to come)

### iCal (OOo only)
(to come)

### Chat (OOo only)
(to come)

### asana connect (OOo only; experimental)
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
- When syncing into unformatted cells, the format might be the (unexpected) default. For instance, dates might show as numbers (which you can change by changing the cell format).

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


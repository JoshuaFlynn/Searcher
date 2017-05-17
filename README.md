# Searcher
Excel spreadsheet that can search a large number of spreadsheets for specific text

Set up the runtime libraries for Searcher as appropriate to your work environment:

Press Alt + F11 to open the VBA menu.

Click Tools, then select References.

In the menu, scroll until you find the following and tick them [note, it may be possible to use Searcher with similar libraries, or without these, depending. These are just the libraries it has by default. Your mileage may vary.]:

Visual Basic For Applications [default]

Microsoft Excel 14.0 Object library [earlier versions may work, needed to open Excel files xlsx/xlsm etc]

OLE Automation [appears to be unneeded, but your mileage may vary.]

Microsoft Office 14.0 Object Library [earlier versions may work, needed for general object manipulations]

Microsoft Forms 2.0 Object Library [needed for the main searcher form to load]

Microsoft Scripting Runtime [needed to use script objects for drive name conversions etc]

- Credit to JohnEffland for pointing out I hadn't included the libraries used

Searcher V5

How to use Searcher:

Download the spreadsheet to any machine that supports Excel 2010 or later (Searcher may work on earlier versions of Excel, however it has not been tested on earlier versions, so use at own risk). It can be placed anywhere, but for convenience it's recommended to place it in the same directory as the collection of spreadsheets you want to search.

Open up Searcher V5, and enable both editing and macros. The spreadsheet will be blank as Searcher makes use of a UserForm that auto-loads upon opening the file.

You will be presented with an array of options to choose from.


Text to search for at present (as of V5) only supports a maximum of two search terms. Future plans will include replacing it with a powerful logical syntax search engine, but at the meantime it only supports two text terms. You only need one term minimum.

You can specify the logic of the search. The options are:
OR: Find spreadsheets that have either or both options
XOR: Find spreadsheets that only have one term or the other
AND: Find spreadsheets that contain both terms
NAND: Find spreadsheets that *don't* contain both terms.
NOR AKA NEITHER: Find spreadsheets that *don't* contain either term.
BUT NOT: Find spreadsheets that contain Term 1 BUT NOT Term 2.

Next to the above options will be checkboxes saying:
Part?: This will search for the term in part of another word.
Case sensitive?: This will search verbatim what you write, otherwise it will be case insensitive.


Directory to search.
You will then have a specify the directory you want to search, you can:
1) Type it in manually
2) Use the current directory Searcher is in (which is why placing it with the files is a good idea. Searcher automatically excludes itself from search results).
3) Search (navigate to) the directory you want to search.


Search sub-folders.
Searcher has the ability to recursively search subfolders, which you can enable. This will carry the relative performance hit if enabled, however it's reasonably fast.


Password.
This will not display your password (asterisk protected). Searcher does NOT retain the password (individuals may scour the code to check) beyond it's initial search. The password is exclusively for the spreadsheets you are wanting to search through. However, the password is stored plain text in temporary memory for the duration of the search, and a future plan is to give this a basic level of encryption. Password isn't necessary, but searcher can't open password protected spreadsheets without it.

Advanced options.
You will have two options.
Break on first result. The program will stop searching as soon as it finds a valid matching spreadsheet.
Auto-open spreadsheet. The program will keep the successful spreadsheet open. Note that in order to avoid any issues, searcher opens all spreadsheets as READ ONLY, so you will not be able to save changes to the spreadsheet. If you want to do this, open it manually.

Output box.
The output box records Searcher's search results and progress, even notifying which spreadsheets it fails to open. This can be exported to a txt file for larger analysis, or copy and pasted from the box directly.

Common problems:

Annoying dialogue prompts.
Although Searcher does it's best to suppress annoying dialogue prompts (for example, it will supply a default one letter password if no password is given to suppress password prompts, and it tries to silence any dialogue queries with events being disabled), it does not always succeed, especially with 'update links' style prompts. The user will have to attend to Searcher, as it's not truly fully automatic. On spreadsheets where there are no issues, Searcher often runs smoothly.

Searcher V6

Searcher Version 6 is now available in the Beta folder, however it is untested, use at own risk.

If exiting first result is enabled, and opening first result is enabled, this adds the ability to automatically open a file in write mode (as opposed to read only mode).

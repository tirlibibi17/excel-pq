# A template for Power Query projects

This includes a few things that can come in handy when working with Power Query.

## A basic configuration parameter framework

Function GetParam("key") returns the associated value.

## A few defined names that refer to the workbook's properties

These are all based on the result of `CELL("filename")` (stored in `fn_cell_filename`

* `fn_filename`: the workbook's name without the full path
* `fn_directory`: if the file is local, it returns the directory path; if the file is on a Sharepoint site (tested on on-prem Sharepoint 2010 with https enabled), `fn_cell_filename` contains the URL of the file (with extra brackets), but `fn_directory` converts it to UNC such as `\\hostname@SSL\DavWWWRoot\sites\path\to\directory`
* `fn_full_path`: full file path

__Note: if you open the file in Protected mode, these names will not update. Also, once you get out of Protected mode, there's a good chance the cells in the configuration table will display `#N/A!`. Just hit F9 to recalculate.__

### NEW! Perform Regular Expression matches with the RegexMatch function

Thanks to Redditor /u/Senipah for the tip. RegexMatch is a generic function that behaves (more or less) like the JS match function.

The RegexMatch function takes 3 arguments:

* string: the text to perform the match on
* pattern: the RegExp pattern to use
* modifiers (optional): any modifiers, e.g. "i" for case-insensitive matching.

If you pass it a pattern with no capturing groups, it will return a [list with one element containing the match](https://i.imgur.com/EZw3PlP.png) if a match is found or [null](https://i.imgur.com/ic6W0vh.png) if there's no match.

If you pass it a pattern with capturing groups, it will return a [list with the first element containing the match](https://i.imgur.com/k0Wj4pT.png) and subsequent elements containing each capturing group.

## Examples of use

### Load a sheet in current workbook:

    let
        Source = Excel.Workbook(File.Contents(GetParam("fullpath")), null, true),
        sheet1_Sheet = Source{[Item="Sheet1",Kind="Sheet"]}[Data]
    in
        sheet1_Sheet

__One caveat: this will read the file from disk, so save first before using.__

### Load from a subfolder of the workbook's path

Define a configuration key called `subfolder` and put `=fn_directory & "\subfolder"` in the value cell. Go to Data / Get & Transform Data / From File / From Folder..., enter a static folder, and then in the M source, replace all occurrences of the folder name with GetParam("subfolder"). The code will continue to work if you move the workbook and the subfolder to a new path.

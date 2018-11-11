# A template for Power Query projects

This includes a few things that can come in handy when working with Power Query

## A basic configuration parameter framework

Function GetParam("key") return the associated value

## A few defined names that refer to the workbook's properties

These are all based on the result of `CELL("filename")` (stored in `fn_cell_filename`

* `fn_filename`: the workbook's name without the full path
* `fn_directory`: if the file is local, it returns the directory path; if the file is on a Sharepoint site (tested on on-prem Sharepoint 2010 with https enabled), `fn_cell_filename` contains the URL of the file (with extra brackets), but `fn_directory` converts it to UNC such as `\\hostname@SSL\DavWWWRoot\sites\path\to\directory`
* `fn_full_path`: full file path

## Examples of use

### Load a sheet in current workbook:

    let
        Source = Excel.Workbook(File.Contents(GetParam("fullpath")), null, true),
        sheet1_Sheet = Source{[Item="Sheet1",Kind="Sheet"]}[Data]
    in
        sheet1_Sheet

One caveat: this will read the file from disk, so save first before using.

### Load from a subfolder of the workbook's path

Define a configuration key called `subfolder` and put `=fn_directory & "\subfolder"` in the value cell. Go to Data / Get & Transform Data / From File / From Folder..., enter a static folder, and then in the M source, replace all occurrences of the folder name with GetParam("subfolder"). The code will continue to work if you move the workbook and the subfolder to a new path.
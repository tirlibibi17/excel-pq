# CherryPick - A function that allows picking cells by address from a worksheet

The function is in the file CherryPick.xlsx. Its signature is the following:

function CherryPick(`tbl` as table, `listOfCells` as list, *optional* `firstCell` as nullable text)

where:

* `tbl` is a table
* `listOfCells` is a list of text values, e.g. {"A1","B12"}
* `firstCell` is the address of the first cell in the table. This is needed because Power Query prunes empty rows and columns at the top and the left of the worksheet. This parameter can be ommitted if the data starts in A1.

Returns a list of the values of each cell.

A couple of invocation examples (with data in A1 and without) are in the file. Error checking is minimal, so use at your own risk.

Note: the **src** directory contains the file used for developing the function.

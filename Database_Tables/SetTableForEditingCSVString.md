# SetTableForEditingCSVString

## Syntax

cDatabaseTables<a id="LST9389C294_01"></a>.SetTableForEditingCSVString

## VB6 Procedure

Function SetTableForEditingCSVString (TableKey
As String, ByRef TableVersion As Integer,
ByRef csvString As String, Optional sepChar
as string = ",") As Integer

## Parameters

TableKey

Input Item: The table key of the table for which data
is requested.

TableVersion

Returned Item: The version number
of the specified table.

csvString

Input Item: A CSV-formatted string containing all the
table data.

sepChar (Optional)

Optional Input Item: The delimiter between data items,
by default ",". This delimiter must match what is used in the
csvString.

## Remarks:

Returns 0 if the function executes correctly, otherwise
returns nonzero.

Reads a table from a CSV-formatted string and adds it
to a stored table list until either the [ApplyEditedTables](ApplyEditedTables.md)
or [CancelTableEditing](CancelTableEditing.md) command is
issued

## Release Notes

Initial release in version 23.0.0

## See Also

[GetTableForDisplayCSVString](GetTableforDisplayCSVString.md)

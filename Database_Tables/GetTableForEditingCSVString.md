# GetTableForEditingCSVString

## Syntax

cDatabaseTables<a id="LST9389C294_01"></a>.GetTableForEditingCSVString

## VB6 Procedure

Function GetTableForEditingCSVString (TableKey As String, GroupName As String, ByRef TableVersion As Integer, ByRef csvString As String, Optional sepChar as string = ",") As Integer

## Parameters

TableKey

Input Item: The table key of the table for which data is requested.

GroupName

Input Item: IMPORTANT NOTE: This parameter is not active in this release. The name of the group for which data will be returned.

TableVersion

Returned Item: The version number of the specified table.  

csvString

Input Item: A CSV-formatted string containing all the table data.

sepChar (Optional)

Optional Input Item: The delimiter between data items, by default ","

## Remarks:

Returns 0 if the function executes correctly, otherwise returns nonzero.

Returns a single table in a string array for interactive editing.

## Release Notes

Initial release in version 23.0.0

## See Also

[SetTableForEditingCSVFile](SetTableForEditingCSVFile.md)

# GetTableForEditingCSVFile

## Syntax

cDatabaseTables<a id="LST9389C294_01"></a>.GetTableForEditingCSVFile

## VB6 Procedure

Function GetTableForEditingCSVFile (TableKey As String(), GroupName As String(), ByRef TableVersion As Integer, csvFilePath As String, Optional sepChar as string = ",") As Integer

## Parameters

TableKey

Input Item: The table key of the table for which data is requested.

GroupName

Input Item: IMPORTANT NOTE: This parameter is not active in this release.The name of the group for which data will be returned.

TableVersion

Returned Item: The version number of the specified table.  

csvFilePath

Input Item: The fully-qualified path for the CSV file containing the table data.

sepChar (Optional)

Optional Input Item: The delimiter between data items, by default ","

## Remarks:

Returns a single table in a CSV file for interactive editing.

Returns a single table in a string array for interactive editing.

## Release Notes

Initial release in version 23.0.0

## See Also

[SetTableForEditingCSVFile](SetTableForEditingCSVFile.md)

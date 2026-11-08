# GetTableForDisplayCSVString

## Syntax

cDatabaseTables<a id="LST9389C294_01"></a>.GetTableForDisplayCSVString

## VB6 Procedure

Function GetTableForDisplayCSVString (TableKey
As String(), ByRef FieldKeyList As
String(), GroupName As String(), ByRef
TableVersion As Integer, csvString
As String(), Optional sepChar as string = ",") As Integer

## Parameters

TableKey

Input Item: The table key of the table for which data
is requested.

FieldKeyList

Input Item: A zero-based array listing the field keys
associated with the specified table for which data is requested.

GroupName

Input Item: The name of the group for which data will
be returned.

TableVersion

Returned Item: The version number
of the specified table.  

csvString

Input Item: A CSV-formatted string containing all the
table data.

sepChar (Optional)

Optional Input Item: The delimiter between data items,
by default ","

## Remarks:

Returns 0 if the function executes correctly, otherwise
returns nonzero.

Returns data for a single table in a CSV string. If
there is nothing to be shown in the table then no data is returned

If the FieldKeyList array
contains a single blank string the data will be provided for all fields.
If the GroupName is All, or a blank string,
then data will be returned for all applicable objects in the model.

## Release Notes

Initial release in version 23.0.0

## See Also

[GetTableForEditingCSVString](GetTableForEditingCSVString.md)

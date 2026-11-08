# GetTableForDisplayXMLString

## Syntax

cDatabaseTables<a id="LST9389C294_01"></a>.GetTableForDisplayXMLString

## VB6 Procedure

Function GetTableForDisplayXMLString (TableKey
As String, ByRef FieldKeyList As String(),
GroupName As String(), IncludeSchema As
Boolean, ByRef TableVersion As Integer,
ByRef XMLTableData As String) As Integer

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

IncludeSchema

Input Item: Flag indicating if the schema should be
included with the table data.

TableVersion

Returned Item: The version number
of the specified table.  

XMLTableData

Returned Item: A string containing the XML data for
the table.

## Remarks:

Returns 0 if the function executes correctly, otherwise
returns nonzero.

Returns data for a single table as XML in a single string.
If there is nothing to be shown in the table then no data is returned

If the FieldKeyList array
contains a single blank string the data will be provided for all fields.
If the GroupName is All, or a blank string,
then data will be returned for all applicable objects in the model.

## Release Notes

Initial release in version 23.0.0

## See Also

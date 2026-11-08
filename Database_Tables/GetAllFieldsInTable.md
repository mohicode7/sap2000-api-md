# GetAllFieldsinTable

## Syntax

cDatabaseTables<a id="LST9389C294_01"></a>.GetAllFieldsInTable

## VB6 Procedure

Function GetAllFieldsInTable (TableKey As String, ByRef
TableVersion As Integer, ByRef NumberFields As Integer, ByRef FieldKey
As String(), FieldName As String(), ByRef Description As String(), ByRef
UnitsString As String(), ByRef IsImportable As Boolean() ) As Integer

## Parameters

TableKey

Input Item: The table key for the table for which the
fields will be returned.

TableVersion

Returned Item: The version number
of the specified table.  

NumberFields

Returned Item: The number of available fields in the
specified table.

FieldKey

Returned Item: A zero-based array of the field keys
for the specified table

FieldName

Returned Item: A zero-based array of the field names
for the specified table

Description

Returned Item: A zero-based array of the field descriptions
for the specified table

UnitsString

Returned Item: A zero-based array of the field units
for the specified table.

IsImportable

Returned Item: A zero-based array of whether the field
is importable for the specified table

## Remarks

Returns the available fields in a specified table.

## Release Notes

Initial release in version 23.0.0

## See Also

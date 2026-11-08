# GetAvailableTables

## Syntax

cDatabaseTables<a id="LST9389C294_01"></a>.GetAvailableTables

## VB6 Procedure

Function GetAvailableTables(ByRef NumberTables As Integer,
ByRef TableKey As As String(), ByRef TableName As As String(), ByRef ImportType
As Integer () ) As Integer

## Parameters

NumberTables

Returned Item: The number of tables that are currently
available for display.

TableKey

Returned Item: zero-based array
of the table keys for the available tables.  

TableName

Returned Item: A zero-based array of the table names
for the available tables.

ImportType

Returned Item: This is either 0, 1, 2 or 3 indicating
the import type for the table.

0= not importable

1= importable, but not interactively
importable

2= importable and interactive importable
when he model is unlocked

3= importable and interactive importable
when he model is unlocked and locked

## Remarks

Returns a list of load cases that are selected for table
display.

## Release Notes

Initial release in version 23.0.0

## See Also

[GetAllTables](GetAllTables.md)

# GetTableForEditingArray

## Syntax

cDatabaseTables<a id="LST9389C294_01"></a>.GetTableForEditingArray

## VB6 Procedure

Function GetTableForEditingArray (TableKey As String(), GroupName As String(), ByRef TableVersion As Integer, ByRef FieldKeysIncluded As String(), ByRef NumberRecords As Integer, ByRef TableData As String()) As Integer

## Parameters

TableKey

Input Item: The table key of the table for which data is requested.

GroupName

Input Item: IMPORTANT NOTE: This parameter is not active in this release.The name of the group for which data will be returned.

TableVersion

Returned Item: The version number of the specified table.  

FieldKeysIncluded

Returned Item: A zero-based array listing the field keys associated with the specified table for which data is reported in the order it is reported in the TableData array. These are essentially the column headers of the data returned in TableData.

NumberRecords

Returned Item: The number of records of data returned for each field. This is essentially the number of rows of data .

TableData

Returned Item: A zero-based, one-dimensional array of the table data, excluding headers, returned row by row. The format of the data is explained below.

As an example, suppose there are three fields in the FieldsKeysIncluded array and the NumberRecords is five.

| EXAMPLE Material Name | Type | Density |
| --- | --- | --- |
| A992Fy50 | Steel | 490 |
| 4000Psi | Concrete | 150 |
| A615Gr60 | Rebar | 480 |
| A416Gr270 | Tendon | 470 |
| 6061T6 | Aluminum | 170 |

The data is returned row by row, and items 0 thru 2 in this array would be the first row, items 3 thru 5 the second row, etc, and items 12 thru 14 the fifth row. Using the example of the table above, the returned array would look like this: (A992Fy50, Steel, 490, 4000Psi, Concrete, 150, A615Gr60, Rebar, 480, A416Gr270, Tendon, 470, 6061T6, Aluminum, 170)

## Remarks:

Returns 0 if the function executes correctly, otherwise returns nonzero.

Returns a single table in a string array for interactive editing.

## Release Notes

Initial release in version 23.0.0

## See Also

[SetTableForEditingArray](SetTableforEditingArray.md)

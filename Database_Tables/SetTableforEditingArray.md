# SetTableForEditingArray

## Syntax

cDatabaseTables<a id="LST9389C294_01"></a>.SetTableForEditingArray

## VB6 Procedure

Function SetTableForEditingArray (TableKey
As String(), ByRef TableVersion As
Integer, ByRef FieldKeysIncluded As String(),
NumberRecords As Integer, ByRef TableData
As String() ) As Integer

## Parameters

TableKey

Input Item: The table key for a table which has been
interactively edited. The table must be one that can be interactively
edited.

TableVersion

Returned Item: The version number
of the specified table.  

FieldKeysIncluded

Input Item: A zero-based array listing the field keys
associated with the specified table for which data is reported in the
order it is reported in the TableData array.
These are essentially the column headers of the data reported in TableData.

NumberRecords

Input Item: The number of records of data returned for
each field. This is essentially the number of rows of data.

Table Data

Input Item: A zero-based, one-dimensional array of the
table data, excluding headers, reported row by row. See [GetTableForDisplayArray](GetTableForEditingArray.md)
for a more detailed explanation of the data format.

## Remarks:

Returns 0 if the function executes correctly, otherwise
returns nonzero.

Reads a table from a string array and adds it to a stored
table list until either the [ApplyEditedTables](ApplyEditedTables.md)
or [CancelTableEditing](CancelTableEditing.md) command is
issued.

Here is an example for how to set the inputs.

TableKey = "EXAMPLE Material Table",

FieldsKeysIncluded =
("Material Name", "Material Type", "Density")

TableData = (A992Fy50, Steel, 490, 4000Psi, Concrete,
150, A615Gr60, Rebar, 480, A416Gr270, Tendon, 470, 6061T6, Aluminum, 170)

Executing the function with these inputs would be equivalent
to setting the following table interactively:

| EXAMPLE Material Name | Type | Density |
| --- | --- | --- |
| A992Fy50 | Steel | 490 |
| 4000Psi | Concrete | 150 |
| A615Gr60 | Rebar | 480 |
| A416Gr270 | Tendon | 470 |
| 6061T6 | Aluminum | 170 |

## Release Notes

Initial release in version 23.0.0

## See Also

[GetTableForDisplayArray](GetTableforDisplayArray.md)

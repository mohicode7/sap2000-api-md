# GetLoadCasesSelectedForDisplay

## Syntax

cDatabaseTables<a id="LST9389C294_01"></a>.GetLoadCasesSelectedForDisplay

## VB6 Procedure

Function GetLoadCasesSelectedForDisplay (ByRef NumberSelectedLoadCases
 As Integer, ByRef LoadCaseList As As String()) As Integer

## Parameters

NumberSelectedLoadCases

Returned Item: The number of load cases selected for
table display.

LoadCaseList

Returned Item: The zero-based
list of load cases selected for table display.  

## Remarks

Returns a list of load cases that are selected for table
display.

This list sets the load cases that are included when
displaying load assignments on the model.

## Release Notes

Initial release in version 23.0.0

## See Also

[SetLoadCasesSelectedForDisplay](SetLoadCasesSelectedForDisplay.md)

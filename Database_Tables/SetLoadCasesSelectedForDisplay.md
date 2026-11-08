# SetLoadCasesSelectedForDisplay

## Syntax

cDatabaseTables<a id="LST9389C294_01"></a>.SetLoadCasesSelectedForDisplay

## VB6 Procedure

Function SetLoadCasesSelectedForDisplay (ByRef LoadCaseList
As As String() ) As Integer

## Parameters

LoadCaseList

Returned Item: The zero-based
list of load cases selected for table display.  

## Remarks

Sets the load cases that are selected for table display.
Returns a 0 if the function executes correctly.

This list sets the load cases that are included when
displaying analysis results. If no load cases are to be selected then
the LoadCaseList item should include a
single blank string.

## Release Notes

Initial release in version 23.0.0

## See Also

[GetLoadCasesSelectedForDisplay](GetLoadCasesSelectedForDisplay.md)

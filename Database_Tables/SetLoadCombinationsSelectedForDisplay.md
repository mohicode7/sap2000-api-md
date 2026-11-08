# SetLoadCombinationsSelectedForDisplay

## Syntax

cDatabaseTables<a id="LST9389C294_01"></a>.SetLoadCombinationsSelectedForDisplay

## VB6 Procedure

Function SetLoadCombinationsSelectedForDisplay (ByRef
LoadCombinationList As As String()) As Integer

## Parameters

LoadCombinationList

Returned Item: The zero-based
list of load combinations selected for table display.  

## Remarks

Sets the load combinations that are selected for table
display. Returns a 0 if the function executes correctly.

This list sets the load combinations that are included
when displaying analysis results. If no load combinations are to be selected
then the LoadCombinationList item should
include a single blank string.

## Release Notes

Initial release in version 23.0.0

## See Also

[GetLoadCombinationsSelectedForDisplay](GetLoadCombinationsSelectedForDisplay.md)

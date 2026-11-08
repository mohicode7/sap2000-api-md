# SetLoadPatternsSelectedForDisplay

## Syntax

cDatabaseTables<a id="LST9389C294_01"></a>.SetLoadPatternsSelectedForDisplay

## VB6 Procedure

Function SetLoad<a id="LST9389C294_011"></a>PatternsSelectedForDisplay
(ByRef Load<a id="LST9389C294_013"></a>PatternList As As String())
As Integer

## Parameters

LoadCombinationList

Returned Item: The zero-based
list of load p<a id="LST9389C294_016"></a>atterns
selected for table display.  

## Remarks

Sets the load patterns that are selected for table display.
Returns a 0 if the function executes correctly.

This list sets the load patterns that are included when
displaying load assignments on the model. If no load patterns are to be
selected then the LoadPatternList item
should include a single blank string.

## Release Notes

Initial release in version 23.0.0

## See Also

[GetLoadPatternsSelectedForDisplay](GetLoadPatternsSelectedForDisplay.md)

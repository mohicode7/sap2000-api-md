# GetLoadPatternsSelectedForDisplay

## Syntax

cDatabaseTables<a id="LST9389C294_01"></a>.GetLoadPatternsSelectedForDisplay

## VB6 Procedure

Function GetLoad<a id="LST9389C294_011"></a>PatternsSelectedForDisplay
(ByRef NumberSelectedLoad<a id="LST9389C294_012"></a>Patterns  As
Integer, ByRef Load<a id="LST9389C294_013"></a>PatternList As As
String()) As Integer

## Parameters

NumberSelectedLoad<a id="LST9389C294_014"></a>Patterns

Returned Item: The number of load p<a id="LST9389C294_015"></a>attern
selected for table display.

LoadPatternList

Returned Item: The zero-based
list of load p<a id="LST9389C294_016"></a>atterns
selected for table display.  

## Remarks

Returns a list of load p<a id="LST9389C294_017"></a>atterns
that are selected for table display.

This list sets the load patterns that are included when
displaying load assignments on the model.

## Release Notes

Initial release in version 23.0.0

## See Also

[SetLoadPatternsSelectedForDisplay](SetLoadPatternsSelectedForDisplay.md)

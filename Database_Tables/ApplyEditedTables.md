# ApplyEditedTables

## Syntax

cDatabaseTables<a id="LST9389C294_01"></a>.ApplyEditedTables

## VB6 Procedure

Function ApplyEditedTables (FillImportLog As Boolean,
ByRef NumFatalErrors As Integer, ByRef NumErrorMsgs As Integer, ByRef
NumWarnMsgs As Integer, ByRef NumInfoMsgs As Integer, ByRef ImportLog
As String ) As Integer

## Parameters

FillImportLog

Whether the ImportLog
string should be filled. Please note that the import log may be very large.

0 = Import log not requested

1 = Import log requested

NumFatalErrors

Returned Item: The number of
fatal errors that occurred during the import

NumErrorMsgs

Returned Item: The number of error messages logged during
the import

NumWarnMsgs

Returned Item: The
number of warning messages logged during the import

NumInfoMsgs

Returned Item: The
number of informational messages logged during the import

ImportLog

Returned Item: A string containing all messages logged
during the import

## Remarks

Instructs the program to interactively import all of
the tables stored in the table list using the SetTableForEditing... functions.

If the model is locked at the time this command is called
then only tables that can be interactively imported when the model is
locked will be imported.

## Release Notes

Initial release in version 23.0.0

## See Also

[SetTableForEditingArray](SetTableforEditingArray.md)

[SetTableForEditingCSVFile](SetTableForEditingCSVFile.md)

[SetTableForEditngCSVString](SetTableForEditingCSVString.md)

# SetStringsExtendedEntityData

## Syntax

SapObject.SapModel.SetStringsExtendedEntityData

## VB6 Procedure

Function SetStringsExtendedEntityData (ByVal AppName As String, ByVal Key As String, ByRef NumberValues As Long, ByRef Values As String()) As Long

## Parameters

AppName

This is an application name of your choice under which your application previously stored some data. It is recommended you choose a unique name to guarantee no other third-party application resets the data that was set by your application. Application names are stored in their original capitalization and character set, but are compared in a case insensitive and culturally invariant manner for retrieval purposes.

Key

This is an entry name under which the data provided in the remaining arguments was stored. This data can later be retrieved, reset, or deleted by providing a valid application name and entry name. Entry names are stored in their original capitalization and character set, but are compared in a case insensitive and culturally invariant manner for retrieval purposes.

NumberValues

The number of strings to store for the given application name and entry name.

Values

A one-dimensional array of strings containing NumberValues strings. These strings replace any strings previously stored for the given application name and entry name. Calling this function with NumberValues equal to zero is equivalent to erasing any previously stored data.

## Remarks

This function can be used to store metadata for a model, or any other data specific to your application.

## Release Notes

Initial release in version 17.2.0.

## See Also

GetKeysWithStringsExtendedEntityData

GetKeyStringsExtendedEntityData

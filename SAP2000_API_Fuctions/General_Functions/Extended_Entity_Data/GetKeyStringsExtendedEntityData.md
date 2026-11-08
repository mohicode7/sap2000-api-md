# GetKeyStringsExtendedEntityData

## Syntax

SapObject.SapModel.GetKeyStringsExtendedEntityData

## VB6 Procedure

Function GetKeyStringsExtendedEntityData (ByVal AppName As String, ByVal Key As String, ByRef NumberValues As Long, ByRef Values As String()) As Long

## Parameters

AppName

This is an application name of your choice under which your application previously stored some data. It is recommended you choose a unique name to guarantee no other third-party application resets the data that was set by your application. Application names are stored in their original capitalization and character set, but are compared in a case insensitive and culturally invariant manner for retrieval purposes.

Key

This is an entry name under which the data provided in the remaining arguments was stored. Entry names are stored in their original capitalization and character set, but are compared in a case insensitive and culturally invariant manner for retrieval purposes.

NumberValues

The number of strings stored for the given application name and entry name.

Values

This is one-dimensional array of string values. The Values array is created as a dynamic, zero-based, array by the API user:

Dim Values() as String

The array is dimensioned to (NumberValues – 1) inside the Sap2000 program, filled with the string values previously stored under the application name AppName and entry name Key, and returned to the API user.

## Release Notes

Initial release in version 17.2.0.

## See Also

GetKeysWithStringsExtendedEntityData

SetKeyStringsExtendedEntityData

# GetKeysWithStringsExtendedEntityData

## Syntax

SapObject.SapModel.GetKeysWithStringsExtendedEntityData

## VB6 Procedure

Function GetKeysWithStringsExtendedEntityData (ByVal AppName As String, ByRef NumberKeys As Long, ByRef Keys() As String) As Long

## Parameters

AppName

This is an application name of your choice under which your application previously stored some data. It is recommended you choose a unique name to guarantee no other third-party application resets the data that was set by your application. Application names are stored in their original capitalization and character set, but are compared in a case insensitive and culturally invariant manner for retrieval purposes.

NumberKeys

The number of different entries for which data was previously stored under the application name AppName.

Key

This is a one-dimensional array of entry names. The Keys array is created as a dynamic, zero-based, array by the API user:

Dim Keys() as String

The array is dimensioned to (NumberKeys – 1) inside the Sap2000 program, filled with the entry names for which data was previously stored under the application name AppName, and returned to the API user.

## Release Notes

Initial release in version 17.2.0.

## See Also

GetKeyStringsExtendedEntityData

SetKeyStringsExtendedEntityData

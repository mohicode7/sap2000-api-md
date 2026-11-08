# GetObsoleteTableKeyList

## Syntax

SapObject.DatabaseTables.GetObsoleteTableKeyList

## VB6 Procedure

Function GetObsoleteTableKeyList(ByRef NumberTableKeys
As Long, ByRef TableKeyList() As String, ByRef NotesList() As String)
As Long

## Parameters

NumberTableKeys

The number
of obsolete table keys for the program.

TableKeyList

A zero-based
array of the obsolete table keys in the program.

NotesList

A zero-based
array of notes associated with each obsolete table key.

## Remarks

This function retrieves a list of obsolete table keys
for the program.

The function returns zero if the list is successfully
retrieved; otherwise it returns nonzero value.

## VBA Example



```
Sub GetObsoleteTableKeys()

   'dimension variables

      Dim SapObject as cOAPI

      Dim SapModel As cSapModel

      Dim ret As Long

      Dim NumberTableKeys As Long

      Dim TableKeyList() As String

      Dim NotesList() As String



   'create Sap2000 object

      Set SapObject = CreateObject("CSI.SAP2000.API.SapObject")



   'start Sap2000 application

      SapObject.ApplicationStart



   'create SapModel object

      Set SapModel = SapObject.SapModel



   'initialize model

      ret = SapModel.InitializeNewModel



   'create a template model

      ret = SapModel.File.New2DFrame(PortalFrame, 
 2, 144, 2, 288)



   'get obsolete table key data

      ret = SapModel.DatabaseTables. GetObsoleteTableKeyList(NumberTableKeys, 
 TableKeyList, NotesList)



   'close Sap2000

      SapObject.ApplicationExit False

      Set SapModel = Nothing

      Set SapObject = Nothing

End Sub
```



## Release Notes

Initial release in version 23.20.

# SetSectionCutsSelectedForDisplay

## Syntax

SapObject.SapModel.DatabaseTables.SetSectionCutsSelectedForDisplay

## VB6 Procedure

Function SetSectionCutsSelectedForDisplay (ByRef SectionCutList()
As String) As Integer

## Parameters

SectionCutList

The zero-based list of section cuts selected for table
display.

## Remarks

This function sets the section cuts that are included
when displaying analysis results. If no section cuts are to be selected
then the list should include a blank string.

The function returns zero if the information is successfully
set, otherwise it returns a nonzero value.

## VBA Example

This example assumes that a file MyModel.sdb exists
and that it contains section cuts named SCut1 and SCut2.



```
Sub SetSectionCutSelectionData()

   'dimension variables

      Dim SapObject as cOAPI

      Dim SapModel As cSapModel

      Dim ret As Long

      Dim FileName As String

      Dim ItemList() As String



   'create Sap2000 object

      Set SapObject = CreateObject("CSI.SAP2000.API.SapObject")



   'start Sap2000 application

      SapObject.ApplicationStart



   'create SapModel object

      Set SapModel = SapObject.SapModel



   'initialize model

      ret = SapModel.InitializeNewModel



   'open an existing file

      FileName = "C:\SapAPI\MyModel.sdb"

      ret = SapModel.File.OpenFile(FileName)



   fill ItemList

      Redim ItemList(1)

      ItemList(0) = SCut1

      ItemList(1) = SCut2



   'set section cuts selected for display in the analysis 
 result tables

      ret = SapModel. DatabaseTables.SetSectionCutsSelectedForDisplay(ItemList)



   'close Sap2000

      SapObject.ApplicationExit False

      Set SapModel = Nothing

      Set SapObject = Nothing

End Sub
```



## Release Notes

Initial release in version 23.00.

## See Also

[GetSectionCutsSelectedForDisplay](GetSectionCutsSelectedForDisplay.md)

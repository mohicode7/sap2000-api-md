# SetGeneralizedDisplacementsSelectedForDisplay

## Syntax

SapObject.SapModel.DatabaseTables.SetGeneralizedDisplacementsSelectedForDisplay

## VB6 Procedure

Function SetGeneralizedDisplacementsSelectedForDisplay
(ByRef GeneralizedDisplacementList() As String) As Integer

## Parameters

GeneralizedDisplacementList

The zero-based list of generalized displacements selected
for table display.

## Remarks

This function sets the generalized displacements that
are included when displaying analysis results. If no generalized displacements
are to be selected then the list should include a blank string.

The function returns zero if the information is successfully
set, otherwise it returns a nonzero value.

## VBA Example

This example assumes that a file MyModel.sdb exists
and that it contains generalized displacements named GDisp1 and GDisp2.



```
Sub SetGeneralizedDisplacementSelectionData()

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

      ItemList(0) = GDisp1

      ItemList(1) = GDisp2



   'set generalized displacements selected for display in 
 the analysis result tables

      ret = SapModel. DatabaseTables.SetGeneralizedDisplacementsSelectedForDisplay(ItemList)



   'close Sap2000

      SapObject.ApplicationExit False

      Set SapModel = Nothing

      Set SapObject = Nothing

End Sub
```



## Release Notes

Initial release in version 23.00.

## See Also

[GetGeneralizedDisplacementsSelectedForDisplay](GetGeneralizedDisplacementsSelectedForDisplay.md)

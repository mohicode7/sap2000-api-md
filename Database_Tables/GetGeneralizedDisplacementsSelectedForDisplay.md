# GetGeneralizedDisplacementsSelectedForDisplay

## Syntax

SapObject.SapModel.DatabaseTables.GetGeneralizedDisplacementsSelectedForDisplay

## VB6 Procedure

Function GetGeneralizedDisplacementsSelectedForDisplay
(ByRef NumberSelectedGeneralizedDisplacements As Integer, ByRef GeneralizedDisplacementList()
As String) As Integer

## Parameters

NumberSelectedGeneralizedDisplacements

The number of generalized displacements selected for
display in the analysis result tables.

GeneralizedDisplacementList

A zero-based list of the names of the generalized displacements
selected for display in the analysis result tables.

## Remarks

This function returns the generalized displacements
selected for display in the analysis result tables.

The function returns zero if the information is successfully
retrieved, otherwise it returns a nonzero value.

## VBA Example

This example assumes that a file MyModel.sdb exists.



```
Sub GetGeneralizedDisplacementSelectionData()

   'dimension variables

      Dim SapObject as cOAPI

      Dim SapModel As cSapModel

      Dim ret As Long

      Dim FileName As String

      Dim NumberItems As Integer

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



   'get generalized displacements selected for display in 
 the analysis result tables

      ret = SapModel. DatabaseTables.GetGeneralizedDisplacementsSelectedForDisplay 
 (NumberItems, ItemList)



   'close Sap2000

      SapObject.ApplicationExit False

      Set SapModel = Nothing

      Set SapObject = Nothing

End Sub
```



## Release Notes

Initial release in version 23.00.

## See Also

[SetGeneralizedDisplacementsSelectedForDisplay](SetGeneralizedDisplacementsSelectedForDisplay.md)

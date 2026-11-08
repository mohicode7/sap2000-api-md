# SetPlotFunctionTracesNamedSetsSelectedForDisplay

## Syntax

SapObject.SapModel.DatabaseTables.SetPlotFunctionTracesNamedSetsSelectedForDisplay

## VB6 Procedure

Function SetPlotFunctionTracesNamedSetsSelectedForDisplay
(ByRef PlotFunctionTracesNamedSetList() As String) As Integer

## Parameters

PlotFunctionTracesNamedSetList

The zero-based list of plot function traces named sets
selected for table display.

## Remarks

This function sets the plot function traces named sets
that are included when displaying analysis results. If no plot function
traces named sets are to be selected then the list should include a blank
string.

The function returns zero if the information is successfully
set, otherwise it returns a nonzero value.

## VBA Example

This example assumes that a file MyModel.sdb exists
and that it contains plot function traces named sets named Set1 and Set2.



```
Sub SetPlotFunctionTracesNamedSetSelectionData()

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

      ItemList(0) = Set1

      ItemList(1) = Set2



   'set plot function traces named sets selected for display 
 in the analysis result tables

      ret = SapModel. DatabaseTables.SetPlotFunctionTracesNamedSetsSelectedForDisplay(ItemList)



   'close Sap2000

      SapObject.ApplicationExit False

      Set SapModel = Nothing

      Set SapObject = Nothing

End Sub
```



## Release Notes

Initial release in version 23.00.

## See Also

[GetPlotFunctionTracesNamedSetsSelectedForDisplay](GetPlotFunctionTracesNamedSetsSelectedForDisplay.md)

# GetCaseSelectedForOutput

## Syntax

SapObject.SapModel.Results.Setup.GetCaseSelectedForOutput

## VB6 Procedure

Function GetCaseSelectedForOutput(ByVal Name As String, ByRef Selected As Boolean) As Long

## Parameters

Name

The name of an existing load case.

Selected

This item is True if the specified load case is selected for output.

## Remarks

This function checks if an load case is selected for output.

The function returns 0 if the selected flag is successfully retrieved, otherwise it returns nonzero.

## VBA Example



```
Sub GetLoadCaseSelected()

   'dimension variables

      Dim SapObject as cOAPI

      Dim SapModel As cSapModel

      Dim ret As Long

      dim Selected As Boolean



   'create Sap2000 object

      Set SapObject = CreateObject("CSI.SAP2000.API.SapObject")



   'start Sap2000 application

      SapObject.ApplicationStart



   'create SapModel object

      Set SapModel = SapObject.SapModel



   'initialize model

      ret = SapModel.InitializeNewModel



   'create model from template

      ret = SapModel.File.New2DFrame(PortalFrame, 3, 124, 3, 200)



   'run analysis

      ret = SapModel.File.Save("C:\SapAPI\x.sdb")

      ret = SapModel.Analyze.RunAnalysis



   'check if case is selected

      ret = SapModel.Results.Setup.GetCaseSelectedForOutput("DEAD", Selected)



   'close Sap2000

      SapObject.ApplicationExit False

      Set SapModel = Nothing

      Set SapObject = Nothing

End Sub
```



## Release Notes

Initial release in version 11.00.

Changed nomenclature from Load Cases, Analysis Cases and Response Combinations to Load Patterns, Load Cases and Load Combinations, respectively, in version 12.00.

## See Also

[DeselectAllCasesAndCombosForOutput](DeselectAllCasesAndCombosForOutput.md)

[GetComboSelectedForOutput](GetComboSelectedForOutput.md)

[SetCaseSelectedForOutput](SetCaseSelectedForOutput.md)

[SetComboSelectedForOutput](SetComboSelectedForOutput.md)

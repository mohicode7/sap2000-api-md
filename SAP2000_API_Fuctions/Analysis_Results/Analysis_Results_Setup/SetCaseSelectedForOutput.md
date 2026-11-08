# SetCaseSelectedForOutput

## Syntax

SapObject.SapModel.Results.Setup.SetCaseSelectedForOutput

## VB6 Procedure

Function SetCaseSelectedForOutput(ByVal Name As String, Optional ByVal Selected As Boolean = True) As Long

## Parameters

Name

The name of an existing load case.

Selected

This item is True if the specified load case is to be selected for output, otherwise it is False.

## Remarks

This function sets an load case selected for output flag.

The function returns 0 if the selected flag is successfully set, otherwise it returns nonzero.

## VBA Example



```
Sub SetLoadCaseSelected()

   'dimension variables

      Dim SapObject as cOAPI

      Dim SapModel As cSapModel

      Dim ret As Long

      Dim Selected As Boolean



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



   'deselect all cases and combos

      ret = SapModel.Results.Setup.DeselectAllCasesAndCombosForOutput



   'set case selected for output

      ret = SapModel.Results.Setup.SetCaseSelectedForOutput("DEAD")



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

[GetCaseSelectedForOutput](GetCaseSelectedForOutput.md)

[GetComboSelectedForOutput](GetComboSelectedForOutput.md)

[SetComboSelectedForOutput](SetComboSelectedForOutput.md)

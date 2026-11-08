# SetComboSelectedForOutput

## Syntax

SapObject.SapModel.Results.Setup.SetComboSelectedForOutput

## VB6 Procedure

Function SetComboSelectedForOutput(ByVal Name As String, Optional ByVal Selected As Boolean = True) As Long

## Parameters

Name

The name of an existing load combination.

Selected

This item is True if the specified load combination is to be selected for output, otherwise it is False.

## Remarks

This function sets a load combination selected for output flag.

The function returns 0 if the selected flag is successfully set, otherwise it returns nonzero.

## VBA Example



```
Sub SetComboSelected()

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



   'open existing model

      ret = SapModel.File.OpenFile("C:\SapAPI\Example 7-002.sdb")



   'run analysis

      ret = SapModel.Analyze.RunAnalysis



   'deselect all cases and combos

      ret = SapModel.Results.Setup.DeselectAllCasesAndCombosForOutput



   'set combo selected for output

      ret = SapModel.Results.Setup.SetComboSelectedForOutput("COMB1")



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

[SetCaseSelectedForOutput](SetCaseSelectedForOutput.md)

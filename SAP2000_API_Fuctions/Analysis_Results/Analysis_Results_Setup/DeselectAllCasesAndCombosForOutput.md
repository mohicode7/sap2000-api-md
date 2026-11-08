# DeselectAllCasesAndCombosForOutput

## Syntax

SapObject.SapModel.Results.Setup.DeselectAllCasesAndCombosForOutput

## VB6 Procedure

Function DeselectAllCasesAndCombosForOutput() As Long

## Parameters

None

## Remarks

The function deselects all load cases and response combinations for output.

The function returns zero if the cases and combos are successfully deselected, otherwise it returns a nonzero value.

## VBA Example



```
Sub DeselectCasesAndCombos()

   'dimension variables

      Dim SapObject as cOAPI

      Dim SapModel As cSapModel

      Dim ret As Long

      Dim NumberResults As Long

      Dim PointElm() As String

      Dim LoadCase() As String

      Dim StepType() As String

      Dim StepNum() As Double

      Dim U1() As Double

      Dim U2() As Double

      Dim U3() As Double

      Dim R1() As Double

      Dim R2() As Double

      Dim R3() As Double



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

[GetCaseSelectedForOutput](GetCaseSelectedForOutput.md)

[GetComboSelectedForOutput](GetComboSelectedForOutput.md)

[SetCaseSelectedForOutput](SetCaseSelectedForOutput.md)

[SetComboSelectedForOutput](SetComboSelectedForOutput.md)

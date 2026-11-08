# BucklingFactor

## Syntax

SapObject.SapModel.Results.BucklingFactor

## VB6 Procedure

Function BucklingFactor(ByRef NumberResults As Long, ByRef LoadCase() As String, ByRef StepType() As String, ByRef stepnum() As Double, ByRef Factor() As Double) As Long

## Parameters

NumberResults

The total number of results returned by the program.

LoadCase

This is an array that includes the name of the analysis case or load combination associated with each result.

StepType

This is an array that includes the step type for each result. For buckling factors, the step type is always Mode. See [Analysis Results Step Type](Analysis_Results_Step_Type.md).

StepNum

This is an array that includes the step number for each result. For buckling factors, the step number is always the buckling mode number. See [Analysis Results Step Number](Analysis_Results_Step_Number.md).

Factor

This is an array that includes the buckling factors.

## Remarks

This function reports buckling factors obtained from buckling load cases.

The function returns zero if the factors are successfully recovered, otherwise it returns a nonzero value.

See [Analysis Results Remarks](Analysis_Results_Remarks.md) for more information.

## VBA Example



```
Sub GetBucklingFactors()

       'dimension variables

            Dim SapObject as cOAPI

          Dim SapModel As cSapModel

          Dim ret As Long

          Dim NumberResults As Long

          Dim LoadCase() As String

          Dim StepType() As String

          Dim StepNum() As Double

          Dim Factor() As Double



       'create Sap2000 object

          Set SapObject = CreateObject("CSI.SAP2000.API.SapObject")



       'start Sap2000 application

          SapObject.ApplicationStart



       'create SapModel object

          Set SapModel = SapObject.SapModel



       'initialize model

          ret = SapModel.InitializeNewModel



       'open model

          ret = SapModel.File.OpenFile("C:\SapAPI\Example 1-019a.sdb")



       'run analysis

          ret = SapModel.Analyze.RunAnalysis



       'clear all case and combo output selections

          ret = SapModel.Results.Setup.DeselectAllCasesAndCombosForOutput



       'set case and combo output selections

          ret = SapModel.Results.Setup.SetCaseSelectedForOutput("BUCK1")



       'get buckling factors

          ret = SapModel.Results.BucklingFactor(NumberResults, LoadCase, StepType, StepNum, Factor)



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

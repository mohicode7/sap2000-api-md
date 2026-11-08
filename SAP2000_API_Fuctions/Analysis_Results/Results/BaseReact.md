# BaseReact

## Syntax

SapObject.SapModel.Results.BaseReact

## VB6 Procedure

Function BaseReact(ByRef NumberResults As Long, ByRef LoadCase() As String, ByRef StepType() As String, ByRef StepNum() As Double, ByRef Fx() As Double, ByRef Fy() As Double, ByRef Fz() As Double, ByRef Mx() As Double, ByRef My() As Double, ByRef Mz() As Double, ByRef gx as Double, ByRef gy as Double, ByRef gz as Double) As Long

## Parameters

NumberResults

The total number of results returned by the program.

LoadCase

This is an array that includes the name of the analysis case or load combination associated with each result.

StepType

This is an array that includes the step type, if any, for each result. See [Analysis Results Step Type](Analysis_Results_Step_Type.md).

StepNum

This is an array that includes the step number, if any, for each result. See [Analysis Results Step Number](Analysis_Results_Step_Number.md).

Fx, Fy, Fz

These are one dimensional arrays that include the base reaction forces in the global X, Y and Z directions, respectively, for each result. [F]

Mx, My, Mz

These are one dimensional arrays that include the base reaction moments about the global X, Y and Z axes, respectively, for each result. [FL]

gx, gy, gz

These are the global X, Y and Z coordinates of the point at which the base reactions are reported. [L]

## Remarks

This function reports the structure total base reactions.

The function returns zero if the reactions are successfully recovered, otherwise it returns a nonzero value.

See [Analysis Results Remarks](Analysis_Results_Remarks.md) for additional information.

## VBA Example



```
Sub GetBaseReactions()

       'dimension variables

          Dim SapObject as cOAPI

          Dim SapModel As cSapModel

          Dim ret As Long

          Dim NumberResults As Long

          Dim PointElm() As String

          Dim LoadCase() As String

          Dim StepType() As String

          Dim StepNum() As Double

          Dim Fx() As Double

          Dim Fy() As Double

          Dim Fz() As Double

          Dim Mx() As Double

          Dim My() As Double

          Dim Mz() As Double

          Dim gx As Double

          Dim gy As Double

          Dim gz As Double



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



       'clear all case and combo output selections

          ret = SapModel.Results.Setup.DeselectAllCasesAndCombosForOutput



       'set case and combo output selections

          ret = SapModel.Results.Setup.SetCaseSelectedForOutput("DEAD")



       'get base reactions

          ret = SapModel.Results.BaseReact(NumberResults, LoadCase, StepType, StepNum, Fx, Fy, Fz, Mx, My, Mz, gx, gy, gz)



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

[BaseReactWithCentroid](BaseReactWithCentroid.md)

[GetJointReact](JointReact.md)

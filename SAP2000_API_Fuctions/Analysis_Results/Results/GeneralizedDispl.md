# GeneralizedDispl

## Syntax

SapObject.SapModel.Results.GeneralizedDispl

## VB6 Procedure

Function GeneralizedDispl(ByVal Name As String, ByRef NumberResults As Long, ByRef GD() As String, ByRef LoadCase() As String, ByRef StepType() As String, ByRef StepNum() As Double, ByRef DType() As String, ByRef Value() As Double) As Long

## Parameters

Name

The name of an existing generalized displacement for which results are returned. If the program does not recognize this name as a defined generalized displacement, it returns results for all selected generalized displacements, if any. For example, entering a blank string (i.e., "") for the name will prompt the program to return results for all selected generalized displacements.

NumberResults

The total number of results returned by the program.

GD

This is an array that includes the generalized displacement name associated with each result.

LoadCase

This is an array that includes the name of the analysis case or load combination associated with each result.

StepType

This is an array that includes the step type, if any, for each result. See [Analysis Results Step Type](Analysis_Results_Step_Type.md).

StepNum

This is an array that includes the step number, if any, for each result. See [Analysis Results Step Number](Analysis_Results_Step_Number.md).

DType

This is an array that includes the generalized displacement type for each result. It is either Translation or Rotation.

Value

This is an array of the generalized displacement values for each result.[L] when DType is Translation , [rad] when DType is Rotation.

## Remarks

This function reports the displacement values for the specified generalized displacements.

The function returns zero if the displacements are successfully recovered, otherwise it returns a nonzero value.

See [Analysis Results Remarks](Analysis_Results_Remarks.md) for more information.

## VBA Example



```
Sub GetGeneralizedDisplacements()

       'dimension variables

            Dim SapObject as cOAPI

          Dim SapModel As cSapModel

          Dim ret As Long

          Dim NumberResults As Long

          Dim GD() As String

          Dim LoadCase() As String

          Dim StepType() As String

          Dim StepNum() As Double

          Dim DType() As String

          Dim Value() As Double

          Dim SF() As Double



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



       'add generalized displacement

          ret = SapModel.GDispl.Add("GD1", 1)



       'add point to generalized displacement

          ReDim SF(5)

          SF(0) = 0.5

          ret = SapModel.GDispl.SetPoint("GD1", "3", SF)



       'get generalized displacement results

          ret = SapModel.Results.GeneralizedDispl("GD1",NumberResults, GD, LoadCase, StepType, StepNum, DType, Value)



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

[JointDispl](JointDispl.md)

[JointDisplAbs](JointDisplAbs.md)

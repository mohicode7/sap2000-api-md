# BridgeSuperCutLongitStress

## Syntax

SapObject.SapModel.Results.BridgeSuperCutLongitStress

## VB6 Procedure

Function BridgeSuperCutLongitStress(ByVal Name As String, ByVal CutIndex As Long, ByVal PointIndex As Long, ByRef NumberResults As Long, ByRef LoadCase() As String, ByRef StepType() As String, ByRef StepNum() As Double, ByRef Stress() As Double) As Long

## Parameters

Name

The name of an existing bridge object.

CutIndex

The index number of section cut in this bridge object. This must be from 0 to Count-1, where Count is the value returned by function SapModel.BridgeObj.CountSuperCut.

PointIndex

The index number of the stress point in this section cut in this bridge object. This must be from 0 to Count-1, where Count is the value returned by function SapModel.BridgeObj.CountSuperCutStressPoint.

NumberResults

The total number of results returned.

LoadCase

This is an array that includes the name of the load case or load combination associated with each result.

StepType

This is an array that includes the step type, if any, for each result. See [Analysis Results Step Type](Analysis_Results_Step_Type.md).

StepNum

This is an array that includes the step number, if any, for each result. See [Analysis Results Step Number](Analysis_Results_Step_Number.md).

Stress

This is an array that includes the longitudinal stress value for each result. [F/L2]

## Remarks

This function returns the longitudinal stresses for multiple cases/combos at a single stress point in a superstructure section cut in a bridge object. Use the functions in SapModel.Results.Setup to control the loads and steps for which results are to be obtained.

The function returns zero if the information is successfully retrieved, otherwise it returns a nonzero value.

## VBA Example

This example assumes that a file MyBridge.sdb exists and has a linked bridge object named BOBJ1 in it. It also assumes a load case named DEAD exists in the model.



```
Sub GetBridgeSuperCutLongitStress()

       'dimension variables

          Dim SapObject as cOAPI

          Dim SapModel As cSapModel

          Dim ret As Long

          Dim NumberResults As Long

          Dim StepType() As String

          Dim StepNum() As Double

          Dim Stress() As Double



       'create Sap2000 object

          Set SapObject = CreateObject("CSI.SAP2000.API.SapObject")



       'start Sap2000 application

          SapObject.ApplicationStart



       'create SapModel object

          Set SapModel = SapObject.SapModel



       'initialize model

          ret = SapModel.InitializeNewModel



       'open an existing file

          FileName = "C:\SapAPI\MyBridge.sdb"

          ret = SapModel.File.OpenFile(FileName)



       'run analysis

          ret = SapModel.Analyze.RunAnalysis



       'clear all case and combo output selections

          ret = SapModel.Results.Setup.DeselectAllCasesAndCombosForOutput



       'set case and combo output selections

          ret = SapModel.Results.Setup.SetCaseSelectedForOutput("DEAD")



       'get bridge section cut results

          ret = SapModel.Results.BridgeSuperCutLongitStress("BOBJ1", 1, 1, NumberResults, LoadCase, StepType, StepNum, Stress)



       'close Sap2000

          SapObject.ApplicationExit False

          Set SapModel = Nothing

          Set SapObject = Nothing

    End Sub
  
```



## Release Notes

Initial release in version 12.00.

## See Also

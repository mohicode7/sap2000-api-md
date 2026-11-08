# AreaJointForceShell

## Syntax

SapObject.Sap2000.Results.AreaJointForceShell

## VB6 Procedure

Function AreaJointForceShell(ByVal Name As String, ByVal ItemTypeElm As eItemTypeElm, ByRef NumberResults As Long, ByRef Obj() As String, ByRef Elm() As String, ByRef PointElm() As String, ByRef LoadCase() As String, ByRef StepType() As String, ByRef StepNum() As Double, ByRef F1() As Double, ByRef F2() As Double, ByRef F3() As Double, ByRef M1() As Double, ByRef M2() As Double, ByRef M3() As Double) As Long

## Parameters

Name

The name of an existing area object, area element or group of objects, depending on the value of the ItemTypeElm item.

ItemTypeElm

This is one of the following items in the eItemTypeElm enumeration:

ObjectElm = 0

Element = 1

GroupElm = 2

SelectionElm = 3

If this item is ObjectElm, the result request is for the area elements corresponding to the area object specified by the Name item.

If this item is Element, the result request is for the area element specified by the Name item.

If this item is GroupElm, the result request is for the area elements corresponding to all area objects included in the group specified by the Name item.

If this item is SelectionElm, the result request is for area elements corresponding to all selected area objects and the Name item is ignored.

NumberResults

The total number of results returned by the program.

Obj

This is an array that includes the area object name associated with each result, if any.

Elm

This is an array that includes the area element name associated with each result.

PointElm

This is an array that includes the point element name associated with each result.

LoadCase

This is an array that includes the name of the analysis case or load combination associated with each result.

StepType

This is an array that includes the step type, if any, for each result. See [Analysis Results Step Type](Analysis_Results_Step_Type.md).

StepNum

This is an array that includes the step number, if any, for each result. See [Analysis Results Step Number](Analysis_Results_Step_Number.md).

F1, F2, F3

These are one dimensional arrays that include the joint force components in the point element local axes directions. [F]

M1, M2, M3

These are one dimensional arrays that include the joint moment components about the point element local axes. [FL]

## Remarks

This function reports the area joint forces for the point elements at each corner of the specified area elements that have shell-type properties (not plane or asolid).

The function returns zero if the forces are successfully recovered, otherwise it returns a nonzero value.

See [Analysis Results Remarks](Analysis_Results_Remarks.md) for more information.

## VBA Example



```
Sub GetAreaJointForceShells()

       'dimension variables

          Dim SapObject as cOAPI

          Dim SapModel As cSapModel

          Dim ret As Long

          Dim NumberResults As Long

          Dim Obj() As String

          Dim Elm() As String

          Dim PointElm() As String

          Dim LoadCase() As String

          Dim StepType() As String

          Dim StepNum() As Double

          Dim F1() As Double

          Dim F2() As Double

          Dim F3() As Double

          Dim M1() As Double

          Dim M2() As Double

          Dim M3() As Double



       'create Sap2000 object

          Set SapObject = CreateObject("CSI.SAP2000.API.SapObject")



       'start Sap2000 application

          SapObject.ApplicationStart



       'create SapModel object

          Set SapModel = SapObject.SapModel



       'initialize model

          ret = SapModel.InitializeNewModel



        'create model from template

          ret = SapModel.File.NewWall(6, 48, 6, 48)



       'run analysis

          ret = SapModel.File.Save("C:\SapAPI\x.sdb")

          ret = SapModel.Analyze.RunAnalysis



       'clear all case and combo output selections

          ret = SapModel.Results.Setup.DeselectAllCasesAndCombosForOutput



       'set case and combo output selections

          ret = SapModel.Results.Setup.SetCaseSelectedForOutput("DEAD")



       'get area joint forces for area object "1"

          ret = SapModel.Results.AreaJointForceShell("1", ObjectElm, NumberResults, Obj, Elm, PointElm, LoadCase, StepType, StepNum, F1, F2, F3, M1, M2, M3)



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

[AreaForceShell](AreaForceShell.md)

[AreaStressShell](AreaStressShell.md)

[AreaStressShellLayered](AreaStressShellLayered.md)

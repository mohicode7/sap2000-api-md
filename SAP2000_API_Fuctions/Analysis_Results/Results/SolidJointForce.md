# SolidJointForce

## Syntax

SapObject.SapModel.Results.SolidJointForce

## VB6 Procedure

Function SolidJointForce(ByVal Name As String, ByVal ItemTypeElm As eItemTypeElm, ByRef NumberResults As Long, ByRef Obj() As String, ByRef Elm() As String, ByRef PointElm() As String, ByRef LoadCase() As String, ByRef StepType() As String, ByRef StepNum() As Double, ByRef F1() As Double, ByRef F2() As Double, ByRef F3() As Double, ByRef M1() As Double, ByRef M2() As Double, ByRef M3() As Double) As Long

## Parameters

Name

The name of an existing solid object, solid element, or group of objects, depending on the value of the ItemTypeElm item.

ItemTypeElm

This is one of the following items in the eItemTypeElm enumeration:

ObjectElm = 0

Element = 1

GroupElm = 2

SelectionElm = 3

If this item is ObjectElm, the result request is for the solid elements corresponding to the solid object specified by the Name item.

If this item is Element, the result request is for the solid element specified by the Name item.

If this item is GroupElm, the result request is for the solid elements corresponding to all solid objects included in the group specified by the Name item.

If this item is SelectionElm, the result request is for solid elements corresponding to all selected solid objects and the Name item is ignored.

NumberResults

The total number of results returned by the program.

Obj

This is an array that includes the solid object name associated with each result, if any.

Elm

This is an array that includes the solid element name associated with each result.

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

This function reports the joint forces for the point elements at each corner of the specified solid elements.

The function returns zero if the forces are successfully recovered, otherwise it returns a nonzero value.

See Analysis Results Remarks for more information.

## VBA Example



```
Sub GetSolidJointForces()

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



       'create a solid model from template

          ret = SapModel.File.NewSolidBlock(20, 50, 20)



       'run analysis

          ret = SapModel.File.Save("C:\SapAPI\x.sdb")

          ret = SapModel.Analyze.RunAnalysis



       'clear all case and combo output selections

          ret = SapModel.Results.Setup.DeselectAllCasesAndCombosForOutput



       'set case and combo output selections

          ret = SapModel.Results.Setup.SetCaseSelectedForOutput("DEAD")



       'get solid joint forces for solid object "1"

          ret = SapModel.Results.SolidJointForce("1", ObjectElm, NumberResults, Obj, Elm, PointElm, LoadCase, StepType, StepNum, F1, F2, F3, M1, M2, M3)



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

[SolidStress](SolidStress.md)

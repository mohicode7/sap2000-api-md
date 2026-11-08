# FrameForce

## Syntax

SapObject.SapModel.Results.FrameForce

## VB6 Procedure

Function FrameForce(ByVal Name As String, ByVal ItemTypeElm As eItemTypeElm, ByRef NumberResults As Long, ByRef Obj() As String, ByRef ObjSta() As Double, ByRef Elm() As String, ByRef ElmSta() As Double, ByRef LoadCase() As String, ByRef StepType() As String, ByRef StepNum() As Double, ByRef P() As Double, ByRef V2() As Double, ByRef V3() As Double, ByRef T() As Double, ByRef M2() As Double, ByRef M3() As Double) As Long

## Parameters

Name

The name of an existing line object, line element or group of objects depending on the value of the ItemTypeElm item.

ItemTypeElm

This is one of the following items in the eItemTypeElm enumeration:

ObjectElm = 0

Element = 1

GroupElm = 2

SelectionElm = 3

If this item is ObjectElm, the result request is for the line elements corresponding to the line object specified by the Name item.

If this item is Element, the result request is for the line element specified by the Name item.

If this item is GroupElm, the result request is for the line elements corresponding to all line objects included in the group specified by the Name item.

If this item is SelectionElm, the result request is for line elements corresponding to all selected line objects and the Name item is ignored.

NumberResults

The total number of results returned by the program.

Obj

This is an array that includes the line object name associated with each result, if any.

ObjSta

This is an array that includes the distance measured from the I-end of the line object to the result location.

Elm

This is an array that includes the line element name associated with each result.

ElmSta

This is an array that includes the distance measured from the I-end of the line element to the result location.

LoadCase

This is an array that includes the name of the analysis case or load combination associated with each result.

StepType

This is an array that includes the step type, if any, for each result. See [Analysis Results Step Type](Analysis_Results_Step_Type.md).

StepNum

This is an array that includes the step number, if any, for each result. See [Analysis Results Step Number](Analysis_Results_Step_Number.md).

P, V2, V3

These are one dimensional arrays that include the axial force, shear force in the local 2 direction, and shear force in the local 3 direction, respectively, for each result. [F]

T, M2, M3

These are one dimensional arrays that include the torsion, moment about the local 2axis, and moment about the local 3-axis, respectively, for each result. [FL]

## Remarks

This function reports the frame forces for the specified line elements.

The function returns zero if the forces are successfully recovered, otherwise it returns a nonzero value.

See [Analysis Results Remarks](Analysis_Results_Remarks.md) for additional information.

## VBA Example



```
Sub GetFrameForces()

       'dimension variables

            Dim SapObject as cOAPI

          Dim SapModel As cSapModel

          Dim ret As Long

          Dim NumberResults As Long

          Dim Obj() As String

          Dim ObjSta() As Double

          Dim Elm() As String

          Dim ElmSta() As Double

          Dim LoadCase() As String

          Dim StepType() As String

          Dim StepNum() As Double

          Dim P() As Double

          Dim V2() As Double

          Dim V3() As Double

          Dim T() As Double

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

          ret = SapModel.File.New2DFrame(PortalFrame, 3, 124, 3, 200)



       'run analysis

          ret = SapModel.File.Save("C:\SapAPI\x.sdb")

          ret = SapModel.Analyze.RunAnalysis



       'clear all case and combo output selections

          ret = SapModel.Results.Setup.DeselectAllCasesAndCombosForOutput



       'set case and combo output selections

          ret = SapModel.Results.Setup.SetCaseSelectedForOutput("DEAD")



       'get frame forces for line object "1"

          ret = SapModel.Results.FrameForce("1", Object, NumberResults, Obj, ObjSta, Elm, ElmSta, LoadCase, StepType, StepNum, P, V2, V3, T, M2, M3)



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

[FrameJointForce](FrameJointForce.md)

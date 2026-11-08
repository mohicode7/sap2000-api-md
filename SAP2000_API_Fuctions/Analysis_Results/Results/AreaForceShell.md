# AreaForceShell

## Syntax

SapObject.Sap2000.Results.AreaForceShell

## VB6 Procedure

Function AreaForceShell(ByVal name As String, ByVal ItemTypeElm As eItemTypeElm, ByRef NumberResults As Long, ByRef Obj() As String, ByRef Elm() As String, ByRef PointElm() As String, ByRef LoadCase() As String, ByRef StepType() As String, ByRef StepNum() As Double, ByRef F11() As Double, ByRef F22() As Double, ByRef F12() As Double, ByRef FMax() As Double, ByRef FMin() As Double, ByRef FAngle() As Double, ByRef FVM() As Double, ByRef M11() As Double, ByRef M22() As Double, ByRef M12() As Double, ByRef MMax() As Double, ByRef MMin() As Double, ByRef MAngle() As Double, ByRef V13() As Double, ByRef V23() As Double, ByRef VMax() As Double, ByRef VAngle() As Double) As Long

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

If this item is SelectionElm, the result request is for area elements corresponding to all selected area objects, and the Name item is ignored.

NumberResults

The total number of results returned by the program.

Obj

This is an array that includes the area object name associated with each result, if any.

Elm

This is an array that includes the area element name associated with each result.

PointElm

This is an array that includes the name of the point element where the results are reported.

LoadCase

This is an array that includes the name of the analysis case or load combination associated with each result.

StepType

This is an array that includes the step type, if any, for each result. See [Analysis Results Step Type](Analysis_Results_Step_Type.md).

StepNum

This is an array that includes the step number, if any, for each result. See [Analysis Results Step Number](Analysis_Results_Step_Number.md).

F11

The area element internal F11 membrane direct force per length reported in the area element local coordinate system. [F/L]

F22

The area element internal F22 membrane direct force per length reported in the area element local coordinate system. [F/L]

F12

The area element internal F12 membrane shear force per length reported in the area element local coordinate system. [F/L]

FMax

The maximum principal membrane force per length. [F/L]

FMin

The minimum principal membrane force per length. [F/L]

FAngle

The angle measured counter clockwise (when the local 3 axis is pointing toward you) from the area local 1 axis to the direction of the maximum principal membrane force. [deg]

FVM

The area element internal Von Mises membrane force per length. [F/L]

M11

The area element internal M11 plate bending moment per length reported in the area element local coordinate system. This item is only reported for area elements with properties that allow plate bending behavior. [FL/L]

M22

The area element internal M22 plate bending moment per length reported in the area element local coordinate system. This item is only reported for area elements with properties that allow plate bending behavior. [FL/L]

M12

The area element internal M12 plate twisting moment per length reported in the area element local coordinate system. This item is only reported for area elements with properties that allow plate bending behavior. [FL/L]

MMax

The maximum principal plate moment per length. This item is only reported for area elements with properties that allow plate bending behavior. [FL/L]

MMin

The minimum principal plate moment per length. This item is only reported for area elements with properties that allow plate bending behavior. [FL/L]

MAngle

The angle measured counter clockwise (when the local 3 axis is pointing toward you) from the area local 1 axis to the direction of the maximum principal plate moment. This item is only reported for area elements with properties that allow plate bending behavior. [deg]

V13

The area element internal V13 plate transverse shear force per length reported in the area element local coordinate system. This item is only reported for area elements with properties that allow plate bending behavior. [F/L]

V23

The area element internal V23 plate transverse shear force per length reported in the area element local coordinate system. This item is only reported for area elements with properties that allow plate bending behavior. [F/L]

VMax

The maximum plate transverse shear force. It is equal to the square root of the sum of the squares of V13 and V23. This item is only reported for area elements with properties that allow plate bending behavior. [F/L]

VAngle

The angle measured counter clockwise (when the local 3 axis is pointing toward you) from the area local 1 axis to the direction of Vmax. This item is only reported for area elements with properties that allow plate bending behavior. [deg]

## Remarks

This function reports the area forces for the specified area elements that are assigned shell section properties (not plane or asolid properties). Note that the forces reported are per unit of in-plane length.

The function returns zero if the forces are successfully recovered, otherwise it returns a nonzero value.

See [Analysis Results Remarks](Analysis_Results_Remarks.md) for more information.

## VBA Example



```
Sub GetAreaForces()

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

          Dim F11() As Double

          Dim F22() As Double

          Dim F12() As Double

          Dim FMax() As Double

          Dim FMin() As Double

          Dim FAngle() As Double

          Dim FVM() As Double

          Dim M11() As Double

          Dim M22() As Double

          Dim M12() As Double

          Dim MMax() As Double

          Dim MMin() As Double

          Dim MAngle() As Double

          Dim V13() As Double

          Dim V23() As Double

          Dim VMax() As Double

          Dim VAngle() As Double



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



       'get area forces for area object "1"

          ret = SapModel.Results.AreaForceShell("1", Object, NumberResults, Obj, Elm, PointElm, LoadCase, StepType, StepNum, F11, F22, F12, FMax, FMin, FAngle, FVM, M11, M22, M12, MMax, MMin, MAngle, V13, V23, VMax, VAngle)



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

[AreaStressShell](AreaStressShell.md)

[AreaStressShellLayered](AreaStressShellLayered.md)

[AreaJointForcePlane](AreaJointForcePlane.md)

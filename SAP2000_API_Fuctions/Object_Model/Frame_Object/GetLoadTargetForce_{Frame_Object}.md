# GetLoadTargetForce

## Syntax

SapObject.SapModel.FrameObj.GetLoadTargetForce

## VB6 Procedure

Function GetLoadTargetForce(ByVal Name As String, ByRef NumberItems As Long, ByRef FrameName() As String, ByRef LoadPat() As String, ByRef dof1() As Boolean, ByRef dof2() As Boolean, ByRef dof3() As Boolean, ByRef dof4() As Boolean, ByRef dof5() As Boolean, ByRef dof6() As Boolean, ByRef P() As Double, ByRef V2() As Double, ByRef V3() As Double, ByRef T() As Double, ByRef M2() As Double, ByRef M3() As Double, ByRef T1() As Double, ByRef T2() As Double, ByRef T3() As Double, ByRef T4() As Double, ByRef T5() As Double, ByRef T6() As Double, Optional ByVal ItemType As eItemType = Object) As Long

## Parameters

Name

The name of an existing frame object or group, depending on the value of the ItemType item.

NumberItems

The total number of deformation loads retrieved for the specified frame objects.

FrameName

This is an array that includes the name of the frame object associated with each target force.

LoadPat

This is an array that includes the name of the load pattern associated with each target force.

dof1, dof2, dof3, dof4, dof5, dof6

These are arrays of boolean values indicating if the considered degree of freedom has a target force assignment.

dof1 = P

dof2 = V2

dof3 = V3

dof4 = T

dof5 = M2

dof6 = M3

P, V2, V3, T, M2, M3

These are arrays of target force values. The target forces specified for a given degree of freedom are applicable only if the corresponding DOF item for that degree of freedom is True.

U1 = U1 deformation [L]

U2 = U2 deformation [L]

U3 = U3 deformation [L]

R1 = R1 deformation [rad]

R2 = R2 deformation [rad]

R3 = R3 deformation [rad]

T1, T2, T3, T4, T5, T6

These are arrays of the relative distances along the frame objects where the target force values apply. The relative distances specified for a given degree of freedom are applicable only if the corresponding dofn item for that degree of freedom is True.

T1 = relative location for P target force

T2 = relative location for V2 target force

T3 = relative location for V3 target force

T4 = relative location for T target force

T5 = relative location for M2 target force

T6 = relative location for M3 target force

ItemType

This is one of the following items in the eItemType enumeration:

Object = 0

Group = 1

SelectedObjects = 2

If this item is Object, the assignments are retrieved for the frame object specified by the Name item.

If this item is Group, the assignments are retrieved for all frame objects in the group specified by the Name item.

If this item is SelectedObjects, assignments are retrieved for all selected frame objects, and the Name item is ignored.

## Remarks

This function retrieves the target force assignments to frame objects.

The function returns zero if the target force assignments are successfully retrieved, otherwise it returns a nonzero value.

## VBA Example



```
Sub GetFrameTargetForce()

       'dimension variables

          Dim SapObject as cOAPI

          Dim SapModel As cSapModel

          Dim ret As Long

          Dim DOF() As Boolean

          Dim f() As double

          Dim RD() As double

          Dim NumberItems As Long

          Dim FrameName() As String

          Dim LoadPat() As String

          Dim dof1() As Boolean

          Dim dof2() As Boolean

          Dim dof3() As Boolean

          Dim dof4() As Boolean

          Dim dof5() As Boolean

          Dim dof6() As Boolean

          Dim P() As Double

          Dim V2() As Double

          Dim V3() As Double

          Dim T() As Double

          Dim M2() As Double

          Dim M3() As Double

          Dim T1() As Double

          Dim T2() As Double

          Dim T3() As Double

          Dim T4() As Double

          Dim T5() As Double

          Dim T6() As Double



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



       'assign frame target force

          ReDim DOF(5)

          ReDim f(5)

          ReDim RD(5)

          DOF(0) = True

          f(0) = 50

          RD(0) = 0.5

          ret = SapModel.FrameObj.SetLoadTargetForce("1", "DEAD", DOF, f, RD)



       'get frame target force

          ret = SapModel.FrameObj.GetLoadTargetForce("1", NumberItems, FrameName, LoadPat, dof1, dof2, dof3, dof4, dof5, dof6, P, V2, V3, T, M2, M3, T1, T2, T3, T4, T5, T6)



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

[SetLoadTargetForce](SetLoadTargetForce_%7BFrame_Object%7D.md)

[DeleteLoadTargetForce](DeleteLoadTargetForce_%7BFrame_Object%7D.md)

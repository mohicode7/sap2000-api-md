# GetLoadDeformation

## Syntax

SapObject.SapModel.FrameObj.GetLoadDeformation

## VB6 Procedure

Function GetLoadDeformation(ByVal Name As String, ByRef NumberItems As Long, ByRef FrameName() As String, ByRef LoadPat() As String, ByRef dof1() As Boolean, ByRef dof2() As Boolean, ByRef dof3() As Boolean, ByRef dof4() As Boolean, ByRef dof5() As Boolean, ByRef dof6() As Boolean, ByRef U1() As Double, ByRef U2() As Double, ByRef U3() As Double, ByRef R1() As Double, ByRef R2() As Double, ByRef R3() As Double, Optional ByVal ItemType As eItemType = Object) As Long

## Parameters

Name

The name of an existing frame object or group, depending on the value of the ItemType item.

NumberItems

The total number of deformation loads retrieved for the specified frame objects.

FrameName

This is an array that includes the name of the frame object associated with each deformation load.

LoadPat

This is an array that includes the name of the load pattern associated with each deformation load.

dof1, dof2, dof3, dof4, dof5, dof6

These are arrays of boolean values indicating if the considered degree of freedom has a deformation load.

dof1 = U1

dof2 = U2

dof3 = U3

dof4 = R1

dof5 = R2

dof6 = R3

U1, U2, U3, R1, R2, R3

These are arrays of deformation load values. The deformations specified for a given degree of freedom are applicable only if the corresponding DOF item for that degree of freedom is True.

U1 = U1 deformation [L]

U2 = U2 deformation [L]

U3 = U3 deformation [L]

R1 = R1 deformation [rad]

R2 = R2 deformation [rad]

R3 = R3 deformation [rad]

ItemType

This is one of the following items in the eItemType enumeration:

Object = 0

Group = 1

SelectedObjects = 2

If this item is Object, the assignments are retrieved for the frame object specified by the Name item.

If this item is Group, the assignments are retrieved for all frame objects in the group specified by the Name item.

If this item is SelectedObjects, assignments are retrieved for all selected frame objects, and the Name item is ignored.

## Remarks

This function retrieves the deformation load assignments to frame objects.

The function returns zero if the load assignments are successfully retrieved, otherwise it returns a nonzero value.

## VBA Example



```
Sub GetFrameDeformationLoad()

   'dimension variables

      Dim SapObject as cOAPI

      Dim SapModel As cSapModel

      Dim ret As Long

      Dim DOF() As Boolean

      Dim d() As double

      Dim NumberItems As Long

      Dim FrameName() As String

      Dim LoadPat() As String

      Dim dof1() As Boolean

      Dim dof2() As Boolean

      Dim dof3() As Boolean

      Dim dof4() As Boolean

      Dim dof5() As Boolean

      Dim dof6() As Boolean

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



   'assign frame deformation loads

      ReDim DOF(5)

      ReDim d(5)

      DOF(0) = True

      D(0) = 2

      ret = SapModel.FrameObj.SetLoadDeformation("ALL", "DEAD", DOF, d, Group)



   'get frame deformation loads

      ret = SapModel.FrameObj.GetLoadDeformation("3", NumberItems, FrameName, LoadPat, dof1, dof2, dof3, dof4, dof5, dof6, U1, U2, U3, R1, R2, R3)



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

[SetLoadDeformation](SetLoadDeformation_%7BFrame_Object%7D.md)

[DeleteLoadDeformation](DeleteLoadDeformation_%7BFrame_Object%7D.md)

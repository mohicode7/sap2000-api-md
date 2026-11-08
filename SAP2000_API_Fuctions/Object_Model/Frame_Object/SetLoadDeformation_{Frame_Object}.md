# SetLoadDeformation

## Syntax

SapObject.SapModel.FrameObj.SetLoadDeformation

## VB6 Procedure

Function SetLoadDeformation(ByVal Name As String, ByVal LoadPat As String, ByRef DOF() As Boolean, ByRef d() As Double, Optional ByVal ItemType As eItemType = Object) As Long

## Parameters

Name

The name of an existing frame object or group, depending on the value of the ItemType item.

LoadPat

The name of a defined load pattern.

DOF

This is a array of boolean values indicating if the considered degree of freedom has a deformation load.

DOF(1) = U1

DOF(2) = U2

DOF(3) = U3

DOF(4) = R1

DOF(5) = R2

DOF(6) = R3

d

This is a array of deformation load values. The deformations specified for a given degree of freedom are applied only if the corresponding DOF item for that degree of freedom is True.

d(1) = U1 deformation [L]

d(2) = U2 deformation [L]

d(3) = U3 deformation [L]

d(4) = R1 deformation [rad]

d(5) = R2 deformation [rad]

d(6) = R3 deformation [rad]

ItemType

This is one of the following items in the eItemType enumeration:

Object = 0

Group = 1

SelectedObjects = 2

If this item is Object, the assignment is made to the frame object specified by the Name item.

If this item is Group, the assignment is made to all frame objects in the group specified by the Name item.

If this item is SelectedObjects, assignment is made to all selected frame objects and the Name item is ignored.

## Remarks

This function assigns deformation loads to frame objects.

The function returns zero if the loads are successfully assigned, otherwise it returns a nonzero value.

## VBA Example



```
Sub AssignFrameDeformationLoad()

   'dimension variables

      Dim SapObject as cOAPI

      Dim SapModel As cSapModel

      Dim ret As Long

      Dim DOF() As Boolean

      Dim d() As double



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

[GetLoadDeformation](GetLoadDeformation_%7BFrame_Object%7D.md)

[DeleteLoadDeformation](DeleteLoadDeformation_%7BFrame_Object%7D.md)

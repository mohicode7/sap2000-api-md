# SetLoadTargetForce

## Syntax

SapObject.SapModel.LinkObj.SetLoadTargetForce

## VB6 Procedure

Function SetLoadTargetForce(ByVal Name As String, ByVal LoadPat As String, ByRef DOF() As Boolean, ByRef f() As Double, ByRef RD() As Double, Optional ByVal ItemType As eItemType = Object) As Long

## Parameters

Name

The name of an existing link object or group, depending on the value of the ItemType item.

LoadPat

The name of a defined load pattern.

DOF

This is a array of boolean values indicating if the considered degree of freedom has a target force.

DOF(1) = P

DOF(2) = V2

DOF(3) = V3

DOF(4) = T

DOF(5) = M2

DOF(6) = M3

f

This is a array of target force values. The target forces specified for a given degree of freedom are applied only if the corresponding DOF item for that degree of freedom is True.

f(1) = P [F]

f(2) = V2 [F]

f(3) = V3 [F]

f(4) = T [FL]

f(5) = M2 [FL]

f(6) = M3 [FL]

RD

This is a array of relative distances along the link objects where the target force values apply. The relative distances specified for a given degree of freedom are applicable only if the corresponding DOF item for that degree of freedom is True. The relative distance must be between 0 and 1, 0 <= RD <=1.

RD(1) = relative location for P target force

RD(2) = relative location for V2 target force

RD(3) = relative location for V3 target force

RD(4) = relative location for T target force

RD(5) = relative location for M2 target force

RD(6) = relative location for M3 target force

ItemType

This is one of the following items in the eItemType enumeration:

Object = 0

Group = 1

SelectedObjects = 2

If this item is Object, the assignment is made to the link object specified by the Name item.

If this item is Group, the assignment is made to all link objects in the group specified by the Name item.

If this item is SelectedObjects, assignment is made to all selected link objects, and the Name item is ignored.

## Remarks

This function assigns target forces to frame objects.

The function returns zero if the target forces are successfully assigned; otherwise it returns a nonzero value.

## VBA Example



```
Sub AssignLinkTargetForce()

   'dimension variables

      Dim SapObject as cOAPI

      Dim SapModel As cSapModel

      Dim ret As Long

      Dim DOF() As Boolean

      Dim f() As double

      Dim RD() As double

      Dim Name As String



   'create Sap2000 object

      Set SapObject = CreateObject("CSI.SAP2000.API.SapObject")



   'start Sap2000 application

      SapObject.ApplicationStart



   'create SapModel object

      Set SapModel = SapObject.SapModel



   'initialize model

      ret = SapModel.InitializeNewModel



   'create model from template

      ret = SapModel.File.New2DFrame(PortalFrame, 2, 144, 2, 288)



   'add link object by points

      ret = SapModel.LinkObj.AddByPoint("1", "5", Name)



   'assign link target force

      ReDim DOF(5)

      ReDim f(5)

      ReDim RD(5)

      DOF(0) = True

      f(0) = 50

      RD(0) = 1

      ret = SapModel.LinkObj.SetLoadTargetForce(Name, "DEAD", DOF, f, RD)



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

[GetLoadTargetForce](GetLoadTargetForce_%7BLink_Object%7D.md)

[DeleteLoadTargetForce](DeleteLoadTargetForce_%7BLink_Object%7D.md)

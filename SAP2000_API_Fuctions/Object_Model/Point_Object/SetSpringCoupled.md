# SetSpringCoupled

## Syntax

SapObject.SapModel.PointObj.SetSpringCoupled

## VB6 Procedure

Function SetSpringCoupled(ByVal Name As String, ByRef k() As Double, Optional ByVal ItemType As eItemType = object, Optional ByVal IsLocalCSys As Boolean = False, Optional ByVal Replace As Boolean = False) As Long

## Parameters

Name

The name of an existing point object or group depending on the value of the ItemType item.

k

This is an array of twenty one spring stiffness values.

Value(0) = U1U1 [F/L]

Value(1) = U1U2 [F/L]

Value(2) = U2U2 [F/L]

Value(3) = U1U3 [F/L]

Value(4) = U2U3 [F/L]

Value(5) = U3U3 [F/L]

Value(6) = U1R1 [F/rad]

Value(7) = U2R1 [F/rad]

Value(8) = U3R1 [F/rad]

Value(9) = R1R1 [FL/rad]

Value(10) = U1R2 [F/rad]

Value(11) = U2R2 [F/rad]

Value(12) = U3R2 [F/rad]

Value(13) = R1R2 [FL/rad]

Value(14) = R2R2 [FL/rad]

Value(15) = U1R3 [F/rad]

Value(16) = U2R3 [F/rad]

Value(17) = U3R3 [F/rad]

Value(18) = R1R3 [FL/rad]

Value(19) = R2R3 [FL/rad]

Value(20) = R3R3 [FL/rad]

ItemType

This is one of the following items in the eItemType enumeration:

Object = 0

Group = 1

SelectedObjects = 2

If this item is Object, the spring assignment is made to the point object specified by the Name item.

If this item is Group, the spring assignment is made to all point objects in the group specified by the Name item.

If this item is SelectedObjects, the spring assignment is made to all selected point objects and the Name item is ignored.

IsLocalCSys

If this item is True, the specified spring assignments are in the point object local coordinate system. If it is False, the assignments are in the Global coordinate system.

Replace

If this item is True, all existing point spring assignments to the specified point object(s) are deleted prior to making the assignment. If it is False, the spring assignments are added to any existing assignments.

## Remarks

This function assigns coupled springs to a point object.

The function returns zero if the stiffnesses are successfully assigned, otherwise it returns a nonzero value.

## VBA Example



```
Sub AssignCoupledSpring()

   'dimension variables

      Dim SapObject as cOAPI

      Dim SapModel As cSapModel

      Dim ret As Long

      Dim k() As Double



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



   'assign coupled spring to a point

      ReDim k(20)

      k(2) = 10

      k(17) = 4

      ret = SapModel.PointObj.SetSpringCoupled("3", k)



   'close Sap2000

      SapObject.ApplicationExit False

      Set SapModel = Nothing

      Set SapObject = Nothing

End Sub
```



## Release Notes

Initial release in version 11.00.

## See Also

[GetSpring](GetSpring_%7BPoint_Object%7D.md)

[GetSpringCoupled](GetSpringCoupled_%7BPoint_Object%7D.md)

[SetSpring](SetSpring.md)

[DeleteSpring](DeleteSpring_%7BPoint_Object%7D.md)

[IsSpringCoupled](IsSpringCoupled_%7BPoint_Object%7D.md)

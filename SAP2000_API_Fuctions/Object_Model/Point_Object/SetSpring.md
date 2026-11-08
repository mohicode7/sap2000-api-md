# SetSpring

## Syntax

SapObject.SapModel.PointObj.SetSpring

## VB6 Procedure

Function SetSpring(ByVal Name As String, ByRef k() As Double, Optional ByVal ItemType As eItemType = object, Optional ByVal IsLocalCSys As Boolean = False, Optional ByVal Replace As Boolean = False) As Long

## Parameters

Name

The name of an existing point object or group depending on the value of the ItemType item.

k

This is an array of six spring stiffness values.

Value(0) = U1 [F/L]

Value(1) = U2 [F/L]

Value(2) = U3 [F/L]

Value(3) = R1 [FL/rad]

Value(4) = R2 [FL/rad]

Value(5) = R3 [FL/rad]

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
Sub AssignPointSpring()

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



   'assign spring to a point

      ReDim k(5)

      k(2) = 10

      ret = SapModel.PointObj.SetSpring("3", k)



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

[SetSpringCoupled](GetSpringCoupled_%7BPoint_Object%7D.md)

[DeleteSpring](DeleteSpring_%7BPoint_Object%7D.md)

[IsSpringCoupled](IsSpringCoupled_%7BPoint_Object%7D.md)

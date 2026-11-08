# SetRestraint

## Syntax

SapObject.SapModel.PointObj.SetRestraint

## VB6 Procedure

Function SetRestraint(ByVal Name As String, ByRef Value() As Boolean, Optional ByVal ItemType As eItemType = object) As Long

## Parameters

Name

The name of an existing point object or group depending on the value of the ItemType item.

Value

This is an array of six restraint values.

Value(0) = U1

Value(1) = U2

Value(2) = U3

Value(3) = R1

Value(4) = R2

Value(5) = R3

ItemType

This is one of the following items in the eItemType enumeration:

Object = 0

Group = 1

SelectedObjects = 2

If this item is Object, the restraint assignment is made to the point object specified by the Name item.

If this item is Group, the restraint assignment is made to all point objects in the group specified by the Name item.

If this item is SelectedObjects, the restraint assignment is made to all selected point objects and the Name item is ignored.

## Remarks

This function assigns the restraint assignments for a point object. The restraint assignments are always set in the point local coordinate system.

The function returns zero if the restraint assignments are successfully assigned, otherwise it returns a nonzero value.

## VBA Example



```
Sub AssignPointRestraints()

   'dimension variables

      Dim SapObject as cOAPI

      Dim SapModel As cSapModel

      Dim ret As Long

      Dim i As Long

      Dim Value() As Boolean



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



   'assign point object restraints

      Redim Value(5)

      For i = 0 to 5

         Value(i) = True

      Next i

      ret = SapModel.PointObj.setRestraint("1", Value)



   'close Sap2000

      SapObject.ApplicationExit False

      Set SapModel = Nothing

      Set SapObject = Nothing

End Sub
```



## Release Notes

Initial release in version 11.00.

## See Also

[GetRestraint](GetRestraint_%7BPoint_Object%7D.md)

[DeleteRestraint](DeleteRestraint.md)

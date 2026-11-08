# SetMassByVolume

## Syntax

Sap2000.PointObj.SetMassByVolume

## VB6 Procedure

Function SetMassByVolume(ByVal Name As String, ByVal MatProp As String, ByRef m() As Double, Optional ByVal ItemType As eItemType = Object, Optional ByVal IsLocalCSys As Boolean = True, Optional ByVal Replace As Boolean = False) As Long

## Parameters

Name

The name of an existing point object or group, depending on the value of the ItemType item.

MatProp

The name of an existing material property.

m

This is an array of six mass assignment values.

Value(0) = U1 [L3]

Value(1) = U2 [L3]

Value(2) = U3 [L3]

Value(3) = R1 [L5]

Value(4) = R2 [L5]

Value(5) = R3 [L5]

ItemType

This is one of the following items in the eItemType enumeration:

Object = 0

Group = 1

SelectedObjects = 2

If this item is Object, the mass assignment is made to the point object specified by the Name item.

If this item is Group, the mass assignment is made to all point objects in the group specified by the Name item.

If this item is SelectedObjects, the mass assignment is made to all selected point objects and the Name item is ignored.

IsLocalCSys

If this item is True, the specified mass assignments are in the point object local coordinate system. If it is False, the assignments are in the Global coordinate system.

Replace

If this item is True, all existing point mass assignments to the specified point object(s) are deleted prior to making the assignment. If it is False,  the mass assignments are added to any previously existing assignments.

## Remarks

This function assigns point mass to a point object. The program calculates the mass by multiplying the specified values by the mass per unit volume of the specified material property.

The function returns zero if the mass is successfully assigned; otherwise, it returns a nonzero value.

## VBA Example



```
Sub AssignPointMassByVolume()

   'dimension variables

      Dim SapObject as cOAPI

      Dim ret As Long

      Dim i as long

      Dim m() As Double



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



   'assign point mass

      Redim m(5)

      For i = 0 to 5

         m(i) = 20000 * (i+1)

      Next i

      ret = SapModel.PointObj.SetMassByVolume("3", "4000psi", m)



   'close Sap2000

      SapObject.ApplicationExit False

      Set SapModel= Nothing

      Set SapObject = Nothing

End Sub
```



## Release Notes

Initial release in version 12.00.

## See Also

[SetMass](SetMass.md)

[SetMassByWeight](SetMassByWeight.md)

[GetMass](GetMass_%7BPoint_Object%7D.md)

[DeleteMass](DeleteMass.md)

# SetMass

## Syntax

SapObject.SapModel.PointObj.SetMass

## VB6 Procedure

Function SetMass(ByVal Name As String, ByRef m() As Double, Optional ByVal ItemType As eItemType = object, Optional ByVal IsLocalCSys As Boolean = True, Optional ByVal Replace As Boolean = False) As Long

## Parameters

Name

The name of an existing point object or group depending on the value of the ItemType item.

m

This is an array of six mass assignment values.

Value(0) = U1 [M]

Value(1) = U2 [M]

Value(2) = U3 [M]

Value(3) = R1 [ML2]

Value(4) = R2 [ML2]

Value(5) = R3 [ML2]

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

If this item is True, all existing point mass assignments to the specified point object(s) are deleted prior to making the assignment. If it is False, the mass assignments are added to any existing assignments.

## Remarks

This function assigns point mass to a point object.

The function returns zero if the mass is successfully assigned, otherwise it returns a nonzero value.

## VBA Example



```
Sub AssignPointMass()

   'dimension variables

      Dim SapObject as cOAPI

      Dim SapModel As cSapModel

      Dim ret As Long

      dim i as long

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

         m(i) = (i+1) / 10

      Next i

      ret = SapModel.PointObj.SetMass("3", m)



   'close Sap2000

      SapObject.ApplicationExit False

      Set SapModel = Nothing

      Set SapObject = Nothing

End Sub
```



## Release Notes

Initial release in version 11.00.

## See Also

[GetMass](GetMass_%7BPoint_Object%7D.md)

[DeleteMass](DeleteMass.md)

[SetMassByVolume](SetMassByVolume.md)

[SetMassByWeight](SetMassByWeight.md)

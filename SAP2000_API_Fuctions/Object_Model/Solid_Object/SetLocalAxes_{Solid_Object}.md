# SetLocalAxes

## Syntax

SapObject.SapModel.SolidObj.SetLocalAxes

## VB6 Procedure

Function SetLocalAxes(ByVal Name As String, ByVal a As Double, ByVal b As Double, ByVal c As Double, Optional ByVal ItemType As eItemType = Object) As Long

## Parameters

Name

The name of an existing solid object or group, depending on the value of the ItemType item.

a, b, c

The local axes of the solid object are defined by first setting the positive local 1, 2 and 3 axes the same as the positive global X, Y and Z axes and then doing the following: [deg]

1.    Rotate about the 3 axis by angle a.

2.    Rotate about the resulting 2 axis by angle b.

3.    Rotate about the resulting 1 axis by angle c.

ItemType

This is one of the following items in the eItemType enumeration:

Object = 0

Group = 1

SelectedObjects = 2

If this item is Object, the local axes assignment is made to the solid object specified by the Name item.

If this item is Group, the local axes assignment is made to all solid objects in the group specified by the Name item.

If this item is SelectedObjects, the local axes assignment is made to all selected solid objects, and the Name item is ignored.

## Remarks

This function sets the local axes angles for solid objects.

The function returns zero if the local axes angles are successfully set; otherwise it returns a nonzero value.

## VBA Example



```
Sub SetSolidLocalAxes()

   'dimension variables

      Dim SapObject as cOAPI

      Dim SapModel As cSapModel

      Dim ret As Long



   'create Sap2000 object

      Set SapObject = CreateObject("CSI.SAP2000.API.SapObject")



   'start Sap2000 application

      SapObject.ApplicationStart



   'create SapModel object

      Set SapModel = SapObject.SapModel



   'initialize model

      ret = SapModel.InitializeNewModel



   'create model from template

      ret = SapModel.File.NewSolidBlock(300, 400, 200, , , 2, 2, 2)



   'assign local axes angles

      ret = SapModel.SolidObj.SetLocalAxes("1", 30, 40, 50)



   'close Sap2000

      SapObject.ApplicationExit False

      Set SapModel = Nothing

      Set SapObject = Nothing

End Sub
```



## Release Notes

Initial release in version 11.00.

## See Also

[GetLocalAxes](GetLocalAxes_%7BSolid_Object%7D.md)

# DeleteMass

## Syntax

SapObject.SapModel.AreaObj.DeleteMass

## VB6 Procedure

Function DeleteMass(ByVal Name As String, Optional ByVal ItemType As eItemType = Object) As Long

## Parameters

Name

The name of an existing area object or group, depending on the value of the ItemType item.

ItemType

This is one of the following items in the eItemType enumeration:

Object = 0

Group = 1

SelectedObjects = 2

If this item is Object, the mass assignments are deleted for the area object specified by the Name item.

If this item is Group, the mass assignments are deleted for all area objects in the group specified by the Name item.

If this item is SelectedObjects, the mass assignments are deleted for all selected area objects, and the Name item is ignored.

## Remarks

This function deletes the mass assignments for area objects.

The function returns zero if the mass assignments are successfully deleted; otherwise it returns a nonzero value.

## VBA Example



```
Sub DeleteAreaObjectMass()

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

      ret = SapModel.File.NewWall(2, 48, 2, 48)



   'assign area object mass

      ret = SapModel.AreaObj.SetMass("ALL", .0001, False, Group)



   'delete area object mass

      ret = SapModel.AreaObj.DeleteMass("1")



   'close Sap2000

      SapObject.ApplicationExit False

      Set SapModel = Nothing

      Set SapObject = Nothing

End Sub
```



## Release Notes

Initial release in version 11.00.

## See Also

[GetMass](GetMass_%7BArea_Object%7D.md)

[SetMass](SetMass_%7BArea_Object%7D.md)

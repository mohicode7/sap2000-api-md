# SetProperty

## Syntax

SapObject.SapModel.AreaObj.SetProperty

## VB6 Procedure

Function SetProperty(ByVal name As String, ByVal PropName As String, Optional ByVal ItemType As eItemType = Object) As Long

## Parameters

Name

The name of an existing area object or group, depending on the value of the ItemType item.

PropName

This is None or the name of a area property to be assigned to the specified area object(s). None means that no property is assigned to the area object.

ItemType

This is one of the following items in the eItemType enumeration:

Object = 0

Group = 1

SelectedObjects = 2

If this item is Object, the assignment is made to the area object specified by the Name item.

If this item is Group, the assignment is made to all area objects in the group specified by the Name item.

If this item is SelectedObjects, assignment is made to all selected area objects, and the Name item is ignored.

## Remarks

This function assigns an area property to area objects.

The function returns zero if the property is successfully assigned; otherwise it returns a nonzero value.

## VBA Example



```
Sub SetAreaObjectProp()

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



   'set area property

      ret = SapModel.AreaObj.SetProperty("4", "None")



   'close Sap2000

      SapObject.ApplicationExit False

      Set SapModel = Nothing

      Set SapObject = Nothing

End Sub
```



## Release Notes

Initial release in version 11.00.

## See Also

[GetProperty](GetProperty_%7BArea_Object%7D.md)

# Delete

## Syntax

SapObject.SapModel.SolidObj.Delete

## VB6 Procedure

Function Delete(ByVal Name As String, Optional ByVal ItemType As eItemType = Object) As Long

## Parameters

Name

The name of an existing solid object or group, depending on the value of the ItemType item.

ItemType

This is one of the following items in the eItemType enumeration:

Object = 0

Group = 1

SelectedObjects = 2

If this item is Object, the solid object specified by the Name item is deleted.

If this item is Group, the all solid objects in the group specified by the Name item are deleted.

If this item is SelectedObjects, all selected solid objects are deleted, and the Name item is ignored.

## Remarks

The function deletes solid objects.

The function returns zero if the solid objects are successfully deleted; otherwise it returns a nonzero value.

## VBA Example



```
Sub DeleteSolidObj()

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



   'delete solid object

      ret = SapModel.SolidObj.Delete("2")



   'update view

      ret = SapModel.View.RefreshView(0, True)



   'close Sap2000

      SapObject.ApplicationExit False

      Set SapModel = Nothing

      Set SapObject = Nothing

End Sub
```



## Release Notes

Initial release in version 11.00.

## See Also

[AddByCoord](AddByCoord_%7BSolid_Object%7D.md)

[AddByPoint](AddByPoint_%7BSolid_Object%7D.md)

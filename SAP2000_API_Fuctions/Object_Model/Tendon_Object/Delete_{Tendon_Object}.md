# Delete

## Syntax

SapObject.SapModel.TendonObj.Delete

## VB6 Procedure

Function Delete(ByVal Name As String, Optional ByVal ItemType As eItemType = Object) As Long

## Parameters

Name

The name of an existing tendon object or group depending on the value of the ItemType item.

ItemType

This is one of the following items in the eItemType enumeration:

Object = 0

Group = 1

SelectedObjects = 2

If this item is Object, the tendon object specified by the Name item is deleted.

If this item is Group, all of the tendon objects in the group specified by the Name item are deleted.

If this item is SelectedObjects, all selected tendon objects are deleted, and the Name item is ignored.

## Remarks

The function deletes tendon objects.

The function returns zero if the tendon objects are successfully deleted, otherwise it returns a nonzero value.

## VBA Example



```
Sub DeleteTendonObj()

   'dimension variables

      Dim SapObject as cOAPI

      Dim SapModel As cSapModel

      Dim ret As Long

      Dim Name1 As String

      Dim Name2 As String



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



   'add tendon objects by points

      ret = SapModel.TendonObj.AddByPoint("3", "9", Name1)

      ret = SapModel.TendonObj.AddByPoint("2", "8", Name2)



   'update view

      ret = SapModel.View.RefreshView(0, False)



   'delete tendon object

      ret = SapModel.TendonObj.Delete(Name1)



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

[AddByCoord](AddByCoord_%7BTendon_Object%7D.md)

[AddByPoint](AddByPoint_%7BTendon_Object%7D.md)

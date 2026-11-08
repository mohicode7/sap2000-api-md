# Delete

## Syntax

SapObject.SapModel.LinkObj.Delete

## VB6 Procedure

Function Delete(ByVal Name As String, Optional ByVal ItemType As eItemType = Object) As Long

## Parameters

Name

The name of an existing link object or group, depending on the value of the ItemType item.

ItemType

This is one of the following items in the eItemType enumeration:

Object = 0

Group = 1

SelectedObjects = 2

If this item is Object, the link object specified by the Name item is deleted.

If this item is Group, the all link objects in the group specified by the Name item are deleted.

If this item is SelectedObjects, all selected link objects are deleted, and the Name item is ignored.

## Remarks

The function deletes link objects.

The function returns zero if the link objects are successfully deleted; otherwise it returns a nonzero value.

## VBA Example



```
Sub DeleteLinkObj()

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



   'add link object by points

      ret = SapModel.LinkObj.AddByPoint("3", "", Name1, True)

      ret = SapModel.LinkObj.AddByPoint("1", "5", Name2)



   'delete area object

      ret = SapModel.LinkObj.Delete(Name1)



   'refresh view

      ret = SapModel.View.RefreshView



   'close Sap2000

      SapObject.ApplicationExit False

      Set SapModel = Nothing

      Set SapObject = Nothing

End Sub
```



## Release Notes

Initial release in version 11.00.

## See Also

[AddByCoord](AddByCoord_%7BLink_Object%7D.md)

[AddByPoint](AddByPoint_%7BLink_Object%7D.md)

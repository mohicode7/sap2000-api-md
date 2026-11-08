# Delete

## Syntax

SapObject.SapModel.FrameObj.Delete

## VB6 Procedure

Function Delete(ByVal Name As String, Optional ByVal ItemType As eItemType = Object) As Long

## Parameters

Name

The name of an existing frame object or group depending on the value of the ItemType item.

ItemType

This is one of the following items in the eItemType enumeration:

Object = 0

Group = 1

SelectedObjects = 2

If this item is Object, the frame object specified by the Name item is deleted.

If this item is Group, all of the frame objects in the group specified by the Name item are deleted.

If this item is SelectedObjects, all selected frame objects are deleted, and the Name item is ignored.

## Remarks

The function deletes frame objects.

The function returns zero if the frame object is successfully deleted, otherwise it returns a nonzero value.

## VBA Example



```
Sub DeleteFrameObj()

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

      ret = SapModel.File.New2DFrame(PortalFrame, 3, 124, 3, 200)



   'delete frame object

      ret = SapModel.FrameObj.Delete("1")



   'close Sap2000

      SapObject.ApplicationExit False

      Set SapModel = Nothing

      Set SapObject = Nothing

End Sub
```



## Release Notes

Initial release in version 11.00.

## See Also

[AddByCoord](AddByCoord_%7BFrame_Objects%7D.md)

[AddByPoint](AddByPoint_%7BFrame_Objects%7D.md)

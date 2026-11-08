# DeleteFireproofing

## Syntax

SapObject.SapModel.FrameObj.DeleteFireproofing

## VB6 Procedure

Function DeleteFireproofing(ByVal Name As String, Optional ByVal ItemType As eItemType = Object) As Long

## Parameters

Name

The name of an existing frame object or group depending on the value of the ItemType item.

ItemType

This is one of the following items in the eItemType enumeration:

Object = 0

Group = 1

SelectedObjects = 2

If this item is Object, the fireproofing assignments are deleted for the frame object specified by the Name item.

If this item is Group, the fireproofing assignments are deleted for all frame objects in the group specified by the Name item.

If this item is SelectedObjects, the fireproofing assignments are deleted for all selected frame objects, and the Name item is ignored.

## Remarks

This function deletes the fireproofing assignments for frame objects.

The function returns zero if the fireproofing assignments are successfully deleted, otherwise it returns a nonzero value.

## VBA Example



```
Sub DeleteFireproofing()

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



   'assign fireproofing

      ret = SapModel.FrameObj.SetFireproofing("ALL", 1, 2, 0, 8.68E-06, False, Group)



   'delete fireproofing

      ret = SapModel.FrameObj.DeleteFireproofing("3")



   'close Sap2000

      SapObject.ApplicationExit False

      Set SapModel = Nothing

      Set SapObject = Nothing

End Sub
```



## Release Notes

Initial release in version 11.00.

## See Also

[GetFireproofing](../../Obsolete_Functions/GetFireproofing.md)

[SetFireproofing](../../Obsolete_Functions/SetFireproofing.md)

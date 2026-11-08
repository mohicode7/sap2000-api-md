# DeleteLateralBracing

## Syntax

SapObject.SapModel.FrameObj.DeleteLateralBracing

## VB6 Procedure

Function DeleteLateralBracing(ByVal Name As String, Optional ByVal MyType As Long = 3, Optional ByVal ItemType As eItemType = Object) As Long

## Parameters

Name

The name of an existing frame object or group, depending on the value of the ItemType item.

MyType

This is 1, 2 or 3, indicating the bracing to be deleted.

1 = Delete point bracing

2 = Delete uniform bracing

3 = Delete both point and uniform bracing

ItemType

This is one of the following items in the eItemType enumeration:

Object = 0

Group = 1

SelectedObjects = 2

If this item is Object, the lateral bracing assignments are deleted for the frame object specified by the Name item.

If this item is Group, the lateral bracing assignments are deleted for all frame objects in the group specified by the Name item.

If this item is SelectedObjects then the lateral bracing assignments are deleted for all selected frame objects, and the Name item is ignored.

## Remarks

This function deletes the lateral bracing assignments for frame objects.

The function returns zero if the assignments are successfully deleted; otherwise it returns a nonzero value.

## VBA Example



```
Sub DeleteFrameLatealBracing()

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

      ret = SapModel.File.New2DFrame(PortalFrame, 2, 144, 2, 288)



   'assign frame lateral bracing

      ret = SapModel.FrameObj.SetLateralBracing("8", 1, 3, 0, 0)

      ret = SapModel.FrameObj.SetLateralBracing("8", 1, 3, 0.25, 0)

      ret = SapModel.FrameObj.SetLateralBracing("8", 1, 3, 1, 0)

      ret = SapModel.FrameObj.SetLateralBracing("8", 2, 1, 0.5, 1)



   'delete frame lateral bracing

      ret = SapModel.FrameObj.DeleteLateralBracing("8", 1)



   'close Sap2000

      SapObject.ApplicationExit False

      Set SapModel = Nothing

      Set SapObject = Nothing

End Sub
```



## Release Notes

Initial release in version 11.03.

## See Also

[SetLateralBracing](SetLateralBracing.md)

[GetLateralBracing](GetLateralBracing.md)

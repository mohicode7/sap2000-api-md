# DeleteSpring

## Syntax

SapObject.SapModel.FrameObj.DeleteSpring

## VB6 Procedure

Function DeleteSpring(ByVal Name As String, Optional ByVal ItemType As eItemType = Object) As Long

## Parameters

Name

The name of an existing frame object or group depending on the value of the ItemType item.

ItemType

This is one of the following items in the eItemType enumeration:

Object = 0

Group = 1

SelectedObjects = 2

If this item is Object, the frame spring assignments are deleted for the frame object specified by the Name item.

If this item is Group, the frame spring assignments are deleted for all frame objects in the group specified by the Name item.

If this item is SelectedObjects, the frame spring assignments are deleted for all selected frame objects and the Name item is ignored.

## Remarks

This function deletes all spring assignments for the specified frame objects.

The function returns zero if the assignments are successfully deleted, otherwise it returns a nonzero value.

## VBA Example



```
Sub DeleteFrameSprings()

   'dimension variables

      Dim SapObject as cOAPI

      Dim SapModel As cSapModel

      Dim ret As Long

      Dim Vec() As Double



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



   'assign springs to frame

      ReDim Vec(2)

      ret = SapModel.FrameObj.SetSpring("ALL", 1, 1, 1, "", 1, 2, 0, Vec, 0, False, "Local", Group)



   'delete springs

      ret = SapModel.FrameObj.DeleteSpring("15")



   'close Sap2000

      SapObject.ApplicationExit False

      Set SapModel = Nothing

      Set SapObject = Nothing

End Sub
```



## Release Notes

Initial release in version 11.00.

## See Also

[GetSpring](GetSpring_%7BFrame_Object%7D.md)

[SetSpring](SetSpring__%7BFrame_Object%7D.md)

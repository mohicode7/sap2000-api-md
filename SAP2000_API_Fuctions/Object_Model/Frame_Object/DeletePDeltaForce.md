# DeletePDeltaForce

## Syntax

SapObject.SapModel.FrameObj.DeletePDeltaForce

## VB6 Procedure

Function DeletePDeltaForce(ByVal Name As String, Optional ByVal ItemType As eItemType = Object) As Long

## Parameters

Name

The name of an existing frame object or group, depending on the value of the ItemType item.

ItemType

This is one of the following items in the eItemType enumeration:

Object = 0

Group = 1

SelectedObjects = 2

If this item is Object, the frame modifier assignments are deleted for the frame object specified by the Name item.

If this item is Group, the frame modifier assignments are deleted for all frame objects in the group specified by the Name item.

If this item is SelectedObjects, the frame modifier assignments are deleted for all selected frame objects, and the Name item is ignored.

## Remarks

This function deletes the P-Delta force assignments for frame objects.

The function returns zero if the assignments are successfully deleted, otherwise it returns a nonzero value.

## VBA Example



```
Sub DeletePDeltaForces()

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



   'assign P-Delta force

      ret = SapModel.FrameObj.SetPDeltaForce("ALL", 100, 0, True, , Group)



   'delete P-Delta force

      ret = SapModel.FrameObj.DeletePDeltaForce("3")



   'close Sap2000

      SapObject.ApplicationExit False

      Set SapModel = Nothing

      Set SapObject = Nothing

End Sub
```



## Release Notes

Initial release in version 11.00.

## See Also

[GetPDeltaForce](GetPDeltaForce_%7BFrame_Object%7D.md)

[SetPDeltaForce](SetPDeltaForce.md)

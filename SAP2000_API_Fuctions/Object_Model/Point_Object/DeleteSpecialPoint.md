# DeleteSpecialPoint

## Syntax

SapObject.SapModel.PointObj.DeleteSpecialPoint

## VB6 Procedure

Function DeleteSpecialPoint(ByVal Name As String, Optional ByVal ItemType As eItemType = Object) As Long

## Parameters

Name

The name of an existing point object or group depending on the value of the ItemType item.

ItemType

This is one of the following items in the eItemType enumeration:

Object = 0

Group = 1

SelectedObjects = 2

If this item is Object, the deletion applies to the point object specified by the Name item.

If this item is Group, the deletion applies to all point objects in the group specified by the Name item.

If this item is SelectedObjects, the deletion applies to all selected point objects, and the Name item is ignored.

## Remarks

The function deletes special point objects that have no other objects connected to them.

The function returns zero if the function completes successfully, otherwise it returns a nonzero value.

Point objects can be deleted only if they have no other objects (e.g., frame, cable, tendon, area, solid link) connected to them. If a point object is not specified to be a Special Point, the program automatically deletes that point object when it has no other objects connected to it. If a point object is specified to be a Special Point, to delete it, first delete all other objects connected to the point and then call this function to delete the point.

## VBA Example



```
Sub DeleteSpecialPointObj()

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



   'set as special point

      ret = SapModel.PointObj.SetSpecialPoint("3", True)



   'delete frame objects

      ret = SapModel.FrameObj.Delete("2")

      ret = SapModel.FrameObj.Delete("8")



   'delete special point object

      ret = SapModel.PointObj.DeleteSpecialPoint("3")



   'close Sap2000

      SapObject.ApplicationExit False

      Set SapModel = Nothing

      Set SapObject = Nothing

End Sub
```



## Release Notes

Initial release in version 11.03.

## See Also

[SetSpecialPoint](SetSpecialPoint.md)

[GetSpecialPoint](GetSpecialPoint.md)

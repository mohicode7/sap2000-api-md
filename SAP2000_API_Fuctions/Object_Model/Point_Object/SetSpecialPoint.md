# SetSpecialPoint

## Syntax

SapObject.SapModel.PointObj.SetSpecialPoint

## VB6 Procedure

Function SetSpecialPoint(ByVal Name As String, ByVal SpecialPoint As Boolean, Optional ByVal ItemType As eItemType = Object) As Long

## Parameters

Name

The name of an existing point object or group depending on the value of the ItemType item.

SpecialPoint

This item is True if the point object is specified as a special point, otherwise it is False.

ItemType

This is one of the following items in the eItemType enumeration:

Object = 0

Group = 1

SelectedObjects = 2

If this item is Object, the special point status is set for the point object specified by the Name item.

If this item is Group, the special point status is set for all point objects in the group specified by the Name item.

If this item is SelectedObjects, the special point status is set for all selected point objects and the Name item is ignored.

## Remarks

This function sets the special point status for a point object.

The function returns zero if the special point status is successfully set, otherwise it returns a nonzero value.

Special points are allowed to exist in the model even if no objects (line, area, solid, link) are connected to them. Points that are not special are automatically deleted if no objects connect to them.

## VBA Example



```
Sub SetSpecialPoint()

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



   'set as special point

      ret = SapModel.PointObj.SetSpecialPoint("3", True)



   'close Sap2000

      SapObject.ApplicationExit False

      Set SapModel = Nothing

      Set SapObject = Nothing

End Sub
```



## Release Notes

Initial release in version 11.00.

## See Also

[GetSpecialPoint](GetSpecialPoint.md)

[DeleteSpecialPoint](DeleteSpecialPoint.md)

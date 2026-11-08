# SetSelected

## Syntax

Sap2000.LinkObj.SetSelected

## VB6 Procedure

Function SetSelected(ByVal Name As String, ByVal Selected As Boolean, Optional ByVal ItemType As eItemType = Object) As Long

## Parameters

Name

The name of an existing link object or group, depending on the value of the ItemType item.

Selected

This item is True if the specified link object is selected; otherwise it is False.

ItemType

This is one of the following items in the eItemType enumeration:

Object = 0

Group = 1

SelectedObjects = 2

If this item is Object, the selected status is set for the link object specified by the Name item.

If this item is Group, the selected status is set for all link objects in the group specified by the Name item.

If this item is SelectedObjects, the selected status is set for all selected link objects, and the Name item is ignored.

## Remarks

This function sets the selected status for link objects.

The function returns zero if the selected status is successfully set; otherwise it returns a nonzero value.

## VBA Example



```
Sub SetLinkObjectSelected()

   'dimension variables

      Dim SapObject as cOAPI

      Dim SapModel As cSapModel

      Dim ret As Long

      Dim Name As String



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

      ret = SapModel.LinkObj.AddByPoint("1", "5", Name)



   'set link object selected

      ret = SapModel.LinkObj.SetSelected("ALL", True, Group)



   'update view

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

[GetSelected](GetSelected_%7BLink_Object%7D.md)

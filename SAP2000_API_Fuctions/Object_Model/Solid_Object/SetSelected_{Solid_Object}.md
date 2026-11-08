# SetSelected

## Syntax

Sap2000.SolidObj.SetSelected

## VB6 Procedure

Function SetSelected(ByVal Name As String, ByVal Selected As Boolean, Optional ByVal ItemType As eItemType = Object) As Long

## Parameters

Name

The name of an existing solid object or group, depending on the value of the ItemType item.

Selected

This item is True if the specified solid object is selected, otherwise it is False.

ItemType

This is one of the following items in the eItemType enumeration:

Object = 0

Group = 1

SelectedObjects = 2

If this item is Object, the selected status is set for the solid object specified by the Name item.

If this item is Group, the selected status is set for all solid objects in the group specified by the Name item.

If this item is SelectedObjects, the selected status is set for all selected solid objects, and the Name item is ignored.

## Remarks

This function sets the selected status for solid objects.

The function returns zero if the selected status is successfully set; otherwise it returns a nonzero value.

## VBA Example



```
Sub SetSolidObjectSelected()

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

      ret = SapModel.File.NewSolidBlock(300, 400, 200, , , 2, 2, 2)



   'set solid objects selected

      ret = SapModel.SolidObj.SetSelected("ALL", True, Group)



   'update window

      ret = SapModel.View.RefreshWindow



   'close Sap2000

      SapObject.ApplicationExit False

      Set SapModel = Nothing

      Set SapObject = Nothing

End Sub
```



## Release Notes

Initial release in version 11.00.

## See Also

[GetSelected](GetSelected_%7BSolid_Object%7D.md)

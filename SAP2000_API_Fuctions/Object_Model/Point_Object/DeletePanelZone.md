# DeletePanelZone

## Syntax

SapObject.SapModel.PointObj.DeletePanelZone

## VB6 Procedure

Function DeletePanelZone(ByVal Name As String, Optional ByVal ItemType As eItemType = Object) As Long

## Parameters

Name

The name of a point object or a group depending on the value selected for ItemType item.

ItemType

This is one of the following items from the eItemType enumeration.

Object = 0

Group = 1

SelectedObjects = 2

If Object is selected, the Name item refers to a point object. The panel zone assignments for that point object are removed.

If Group is selected, the Name item refers to a group. The panel zone assignments for all point objects in the group are removed.

If SelectedObjects is selected, the Name item is ignored. The panel zone assignments for all selected point objects are removed.

## Remarks

This function deletes all panel zone assignments from the specified point object(s).

The function returns zero if the panel zone assignments are successfully deleted, otherwise it returns a nonzero value.

## VBA Example



```
Sub DeletePanelZone()

   'dimension variables

      Dim SapObject as cOAPI

      Dim SapModel As cSapModel

      Dim ret as Long

      Dim PropType as Long

      Dim Thickness As Double

      Dim Connectivity As Long

      Dim LocalAxisFrom As Long



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



   'add panel zone assignment

      PropType = 1 'elastic from column and doubler plate

      Thickness = 2

      Connectivity  = 0 'beams to other objects

      LocalAxisFrom = 0 'column

      ret = SapModel.PointObj.SetPanelZone("3", PropType, Thickness, 0, 0, "", Connectivity, LocalAxisFrom, 0)



   'delete panel zone

      ret = SapModel.PointObj.DeletePanelZone("3")



  'close Sap2000

      SapObject.ApplicationExit False

      Set SapModel = Nothing

      Set SapObject = Nothing

End Sub
```



## Release Notes

Initial release in version 11.00.

## See Also

[GetPanelZone](GetPanelZone.md)

[SetPanelZone](SetPanelZone.md)

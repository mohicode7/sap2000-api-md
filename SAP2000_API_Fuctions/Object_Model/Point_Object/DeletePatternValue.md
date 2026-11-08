# DeletePatternValue

## Syntax

SapObject.SapModel.PointObj.DeletePatternValue

## VB6 Procedure

Function DeletePatternValue(ByVal Name As String, ByVal PatternName As String, Optional ByVal ItemType As eItemType = Object) As Long

## Parameters

Name

The name of a point object or a group depending on the value selected for ItemType item.

PatternName

The name of a defined joint pattern.

ItemType

This is one of the following items from the eItemType enumeration.

Object = 0

Group = 1

SelectedObjects = 2

If Object is selected, the Name item refers to a point object. The joint pattern assignments, associated with the specified joint pattern, for that point object are removed.

If Group is selected, the Name item refers to a group. The joint pattern assignments, associated with the specified joint pattern, for all point objects in the group, are removed.

If SelectedObjects is selected, the Name item is ignored. The joint pattern assignments, associated with the specified joint pattern, for all selected point objects, are removed.

## Remarks

This function deletes all joint pattern assignments, associated with the specified joint pattern, from the specified point object(s).

The function returns zero if the joint pattern assignments are successfully deleted, otherwise it returns a nonzero value.

## VBA Example



```
Sub DeletePointPatternAssigns()

   'dimension variables

      Dim SapObject as cOAPI

      Dim SapModel As cSapModel

      Dim ret as Long



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



   'add joint pattern assignment

      ret = SapModel.PointObj.SetPatternByXYZ("3", "Default", 0, 0, 10, 0)



   'delete joint pattern assignment

      ret = SapModel.PointObj.DeletePatternValue("3", "Default")



  'close Sap2000

      SapObject.ApplicationExit False

      Set SapModel = Nothing

      Set SapObject = Nothing

End Sub
```



## Release Notes

Initial release in version 11.00.

## See Also

[GetPatternValue](GetPatternValue_%7BPoint_Object%7D.md)

[SetPatternByPressure](SetPatternByPressure.md)

[SetPatternByXYZ](SetPatternByXYZ.md)

# DeleteLoadPorePressure

## Syntax

SapObject.SapModel.SolidObj.DeleteLoadPorePressure

## VB6 Procedure

Function DeleteLoadPorePressure(ByVal Name As String, ByVal LoadPat As String, Optional ByVal ItemType As eItemType = Object) As Long

## Parameters

Name

The name of an existing solid object or group, depending on the value of the ItemType item.

LoadPat

The name of a defined load pattern.

ItemType

This is one of the following items in the eItemType enumeration:

Object = 0

Group = 1

SelectedObjects = 2

If this item is Object, the load assignments are deleted for the solid object specified by the Name item.

If this item is Group, the load assignments are deleted for all solid objects in the group specified by the Name item.

If this item is SelectedObjects, the load assignments are deleted for all selected solid objects, and the Name item is ignored.

## Remarks

This function deletes the pore pressure load assignments to the specified solid objects for the specified load pattern.

The function returns zero if the load assignments are successfully deleted; otherwise it returns a nonzero value.

## VBA Example



```
Sub DeleteSolidObjectPorePressureLoad()

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



   'assign solid object pore pressure load

      ret = SapModel.SolidObj.SetLoadPorePressure("ALL", "DEAD", .1, , , Group)



   'delete solid object pore pressure load

      ret = SapModel.SolidObj.DeleteLoadPorePressure("1", "DEAD")



   'close Sap2000

      SapObject.ApplicationExit False

      Set SapModel = Nothing

      Set SapObject = Nothing

End Sub
```



## Release Notes

Initial release in version 11.00.

Changed nomenclature from Load Cases, Analysis Cases and Response Combinations to Load Patterns, Load Cases and Load Combinations, respectively, in version 12.00.

## See Also

[GetLoadPorePressure](GetLoadPorePressure_%7BSolid_Object%7D.md)

[SetLoadPorePressure](SetLoadPorePressure_%7BSolid_Object%7D.md)

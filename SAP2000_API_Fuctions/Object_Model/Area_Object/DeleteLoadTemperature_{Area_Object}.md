# DeleteLoadTemperature

## Syntax

SapObject.SapModel.AreaObj.DeleteLoadTemperature

## VB6 Procedure

Function DeleteLoadTemperature(ByVal Name As String, ByVal LoadPat As String, Optional ByVal ItemType As eItemType = Object) As Long

## Parameters

Name

The name of an existing area object or group, depending on the value of the ItemType item.

LoadPat

The name of a defined load pattern.

ItemType

This is one of the following items in the eItemType enumeration:

Object = 0

Group = 1

SelectedObjects = 2

If this item is Object, the load assignments are deleted for the area object specified by the Name item.

If this item is Group, the load assignments are deleted for all area objects in the group specified by the Name item.

If this item is SelectedObjects, the load assignments are deleted for all selected area objects, and the Name item is ignored.

## Remarks

This function deletes the temperature load assignments to the specified area objects for the specified load pattern.

The function returns zero if the load assignments are successfully deleted; otherwise it returns a nonzero value.

## VBA Example



```
Sub DeleteAreaObjectTemperatureLoad()

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

      ret = SapModel.File.NewWall(2, 48, 2, 48)



   'assign area object temperature load

      ret = SapModel.AreaObj.SetLoadTemperature("All", "DEAD", 1, 50, , , Group)



   'delete area object temperature load

      ret = SapModel.AreaObj.DeleteLoadTemperature("3", "DEAD")



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

[GetLoadTemperature](GetLoadTemperature_%7BArea_Object%7D.md)

[SetLoadTemperature](SetLoadTemperature_%7BArea_Object%7D.md)

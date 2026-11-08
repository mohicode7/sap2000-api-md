# DeleteLoadStrain

## Syntax

SapObject.SapModel.SolidObj.DeleteLoadStrain

## VB6 Procedure

Function DeleteLoadStrain(ByVal Name As String, ByVal LoadPat As String, ByVal Component As Long, Optional ByVal ItemType As eItemType = Object) As Long

## Parameters

Name

The name of an existing solid object or group, depending on the value of the ItemType item.

LoadPat

The name of a defined load pattern.

Component

This is 1, 2, 3, 4, 5 or 6, indicating the component for which the strain load is to be deleted.

1 = Strain11

2 = Strain22

3 = Strain33

4 = Strain12

5 = Strain13

6 = Strain23

ItemType

This is one of the following items in the eItemType enumeration:

Object = 0

Group = 1

SelectedObjects = 2

If this item is Object, the load assignments are deleted for the solid object specified by the Name item.

If this item is Group, the load assignments are deleted for all solid objects in the group specified by the Name item.

If this item is SelectedObjects, the load assignments are deleted for all selected solid objects, and the Name item is ignored.

## Remarks

This function deletes the strain load assignments to the specified solid objects, for the specified load pattern, for the specified components.

The function returns zero if the load assignments are successfully deleted; otherwise it returns a nonzero value.

## VBA Example



```
Sub DeleteSolidObjectStrainLoad()

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



   'assign solid object strain load

      ret = SapModel.SolidObj.SetLoadStrain("ALL", "DEAD", 1, 0.001, , , Group)



   'delete solid object strain load

      ret = SapModel.SolidObj.DeleteLoadStrain("1", "DEAD", 1)



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

[GetLoadStrain](GetLoadStrain_%7BSolid_Object%7D.md)

[SetLoadStrain](SetLoadStrain_%7BSolid_Object%7D.md)

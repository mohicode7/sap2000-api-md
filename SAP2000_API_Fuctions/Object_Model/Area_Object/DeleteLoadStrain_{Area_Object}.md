# DeleteLoadStrain

## Syntax

SapObject.SapModel.AreaObj.DeleteLoadStrain

## VB6 Procedure

Function DeleteLoadStrain(ByVal Name As String, ByVal LoadPat As String, ByVal Component As Long, Optional ByVal ItemType As eItemType = Object) As Long

## Parameters

Name

The name of an existing area object or group, depending on the value of the ItemType item.

LoadPat

The name of a defined load pattern.

Component

This is 1, 2, 3, 4, 5, 6, 7, 8, or 9, indicating the component to which the strain load is applied.

1 = Strain11

2 = Strain12

3 = Strain13

4 = Curvature11

5 = Curvature22

6 = Curvature33

7 = Strain13

8 = Strain23

9 = Strain33

ItemType

This is one of the following items in the eItemType enumeration:

Object = 0

Group = 1

SelectedObjects = 2

If this item is Object, the load assignments are deleted for the area object specified by the Name item.

If this item is Group, the load assignments are deleted for all area objects in the group specified by the Name item.

If this item is SelectedObjects, the load assignments are deleted for all selected area objects, and the Name item is ignored.

## Remarks

This function deletes the strain load assignments to the specified area objects, for the specified load pattern, for the specified components.

The function returns zero if the load assignments are successfully deleted; otherwise it returns a nonzero value.

## VBA Example



```
Sub DeleteAreaObjectStrainLoad()

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



   'assign area object strain load

      ret = SapModel.AreaObj.SetLoadStrain("ALL", "DEAD", 1, 0.001, , , Group)



   'delete area object strain load

      ret = SapModel.AreaObj.DeleteLoadStrain("3", "DEAD", 1)



   'close Sap2000

      SapObject.ApplicationExit False

      Set SapModel = Nothing

      Set SapObject = Nothing

End Sub
```



## Release Notes

Initial release in version 11.00.

Changed nomenclature from Load Cases, Analysis Cases and Response Combinations to Load Patterns, Load Cases and Load Combinations, respectively, in version 12.00.

Added Strain33 component in v19.0.0.

## See Also

[GetLoadStrain](GetLoadStrain_%7BArea_Object%7D.md)

[SetLoadStrain](SetLoadStrain_%7BArea_Object%7D.md)

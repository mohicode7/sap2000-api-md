# SetLoadStrain

## Syntax

SapObject.SapModel.AreaObj.SetLoadStrain

## VB6 Procedure

Function SetLoadStrain(ByVal Name As String, ByVal LoadPat As String, ByVal Component As Long, ByVal Value As Double, Optional ByVal Replace As Boolean = True, Optional ByVal PatternName As String = "", Optional ByVal ItemType As eItemType = Object) As Long

## Parameters

Name

The name of an existing area object or group, depending on the value of the ItemType item.

LoadPat

The name of a defined load pattern.

Component

This is 1, 2, 3, 4, 5, 6, 7, 8, or 9, indicating the component to which the strain load is applied.

1 = Strain11

2 = Strain22

3 = Strain12

4 = Curvature11

5 = Curvature22

6 = Curvature12

7 = Strain13

8 = Strain23

9 = Strain33

Value

This is the strain load value. [L/L] for Component = 1, 2, 3, 7, 8, and 9 and [1/L] for Component = 4, 5 and 6

Replace

If this item is True, all previous strain loads, if any, assigned to the specified area object(s), in the specified load pattern, for the specified degree of freedom, are deleted before making the new assignment.

PatternName

This is blank or the name of a defined joint pattern. If it is blank, the strain load for the area object is uniform over the object at the value specified by Value.

If PatternName is the name of a defined joint pattern, the strain load for the area object is based on the specified strain value multiplied by the pattern value at the corner points of the area object.

ItemType

This is one of the following items in the eItemType enumeration:

Object = 0

Group = 1

SelectedObjects = 2

If this item is Object, the assignment is made to the area object specified by the Name item.

If this item is Group, the assignment is made to all area objects in the group specified by the Name item.

If this item is SelectedObjects, assignment is made to all selected area objects, and the Name item is ignored.

## Remarks

This function assigns strain loads to area objects.

The function returns zero if the loads are successfully assigned; otherwise it returns a nonzero value.

## VBA Example



```
Sub AssignAreaObjectStrainLoad()

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

[DeleteLoadStrain](DeleteLoadStrain_%7BArea_Object%7D.md)

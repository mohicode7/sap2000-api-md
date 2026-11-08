# SetLoadStrain

## Syntax

SapObject.SapModel.SolidObj.SetLoadStrain

## VB6 Procedure

Function SetLoadStrain(ByVal Name As String, ByVal LoadPat As String, ByVal Component As Long, ByVal Value As Double, Optional ByVal Replace As Boolean = True, Optional ByVal PatternName As String = "", Optional ByVal ItemType As eItemType = Object) As Long

## Parameters

Name

The name of an existing solid object or group, depending on the value of the ItemType item.

LoadPat

The name of a defined load pattern.

Component

This is 1, 2, 3, 4, 5 or 6, indicating the component to which the strain load is applied.

1 = Strain11

2 = Strain22

3 = Strain33

4 = Strain12

5 = Strain13

6 = Strain23

Value

This is the strain load value. [L/L]

Replace

If this item is True, all previous strain loads, if any, assigned to the specified solid object(s), in the specified load pattern, for the specified degree of freedom, are deleted before making the new assignment.

PatternName

This is blank or the name of a defined joint pattern. If it is blank, the strain load for the solid object is uniform over the object at the value specified by Value.

If PatternName is the name of a defined joint pattern, the strain load for the solid object is based on the specified strain value multiplied by the pattern value at the corner point objects of the solid object.

ItemType

This is one of the following items in the eItemType enumeration:

Object = 0

Group = 1

SelectedObjects = 2

If this item is Object, the assignment is made to the solid object specified by the Name item.

If this item is Group, the assignment is made to all solid objects in the group specified by the Name item.

If this item is SelectedObjects, assignment is made to all selected solid objects, and the Name item is ignored.

## Remarks

This function assigns strain loads to solid objects.

The function returns zero if the loads are successfully assigned; otherwise it returns a nonzero value.

## VBA Example



```
Sub AssignSolidObjectStrainLoad()

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

[DeleteLoadStrain](DeleteLoadStrain_%7BSolid_Object%7D.md)

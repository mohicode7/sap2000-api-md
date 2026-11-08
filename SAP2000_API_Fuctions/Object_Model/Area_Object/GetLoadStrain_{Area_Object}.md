# GetLoadStrain

## Syntax

SapObject.SapModel.AreaObj.GetLoadStrain

## VB6 Procedure

Function GetLoadStrain(ByVal Name As String, ByRef NumberItems As Long, ByRef AreaName() As String, ByRef LoadPat() As String, ByRef Component() As Long, ByRef Value() As Double, ByRef PatternName() As String, Optional ByVal ItemType As eItemType = Object) As Long

## Parameters

Name

The name of an existing area object or group, depending on the value of the ItemType item.

NumberItems

The total number of strain loads retrieved for the specified area objects.

AreaName

This is an array that includes the name of the area object associated with each strain load.

LoadPat

This is an array that includes the name of the load pattern associated with each strain load.

Component

This is an array that includes 1, 2, 3, 4, 5, 6, 7, 8, or 9, indicating the component associated with each strain load.

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

This is an array that includes the strain value. [L/L] for Component = 1, 2, 3, 7, 8, and 9, and [1/L] for Component = 4, 5 and 6

PatternName

This is an array that includes the joint pattern name, if any, used to specify the strain load.

ItemType

This is one of the following items in the eItemType enumeration:

Object = 0

Group = 1

SelectedObjects = 2

If this item is Object, the assignments are retrieved for the area object specified by the Name item.

If this item is Group, the assignments are retrieved for all area objects in the group specified by the Name item.

If this item is SelectedObjects, assignments are retrieved for all selected area objects, and the Name item is ignored.

## Remarks

This function retrieves the strain load assignments to area objects.

The function returns zero if the strain load assignments are successfully retrieved; otherwise it returns a nonzero value.

## VBA Example



```
Sub GetAreaObjectStrainLoad()

   'dimension variables

      Dim SapObject as cOAPI

      Dim SapModel As cSapModel

      Dim ret As Long

      Dim NumberItems As Long

      Dim AreaName() As String

      Dim LoadPat() As String

      Dim Component() As Long

      Dim Value() As Double

      Dim PatternName() As String



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



   'get area object strain load

      ret = SapModel.AreaObj.GetLoadStrain("3", NumberItems, AreaName, LoadPat, Component, Value, PatternName)



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

[SetLoadStrain](SetLoadStrain_%7BArea_Object%7D.md)

[DeleteLoadStrain](DeleteLoadStrain_%7BArea_Object%7D.md)

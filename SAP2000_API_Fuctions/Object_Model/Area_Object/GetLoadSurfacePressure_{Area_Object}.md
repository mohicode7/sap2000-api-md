# GetLoadSurfacePressure

## Syntax

SapObject.SapModel.AreaObj.GetLoadSurfacePressure

## VB6 Procedure

Function GetLoadSurfacePressure(ByVal Name As String, ByRef NumberItems As Long, ByRef AreaName() As String, ByRef LoadPat() As String, ByRef Face() As Long, ByRef Value() As Double, ByRef PatternName() As String, Optional ByVal ItemType As eItemType = Object) As Long

## Parameters

Name

The name of an existing area object or group, depending on the value of the ItemType item.

NumberItems

The total number of surface pressure loads retrieved for the specified area objects.

AreaName

This is an array that includes the name of the area object associated with each surface pressure load.

LoadPat

This is an array that includes the name of the load pattern associated with each surface pressure load.

Face

This is an array that includes either -1, -2 or a nonzero, positive integer, indicating the area object face to which the specified load assignment applies.

-1 = Bottom face

-2 = Top face

>0 = Edge face

Note that edge face n is from area object point n to area object point n + 1. For example, edge face 2 is from area object point 2 to area object point 3.

Value

This is an array that includes the surface pressure load value. [F/L2]

PatternName

This is an array that includes the joint pattern name, if any, used to specify the surface pressure load.

ItemType

This is one of the following items in the eItemType enumeration:

Object = 0

Group = 1

SelectedObjects = 2

If this item is Object, the assignments are retrieved for the area object specified by the Name item.

If this item is Group, the assignments are retrieved for all area objects in the group specified by the Name item.

If this item is SelectedObjects, assignments are retrieved for all selected area objects, and the Name item is ignored.

## Remarks

This function retrieves the surface pressure load assignments to area objects.

The function returns zero if the load assignments are successfully retrieved; otherwise it returns a nonzero value.

## VBA Example



```
Sub GetAreaObjectSurfacePressureLoad()

   'dimension variables

      Dim SapObject as cOAPI

      Dim SapModel As cSapModel

      Dim ret As Long

      Dim NumberItems As Long

      Dim AreaName() As String

      Dim LoadPat() As String

      Dim Face() As Long

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



   'assign area object surface pressure load

      ret = SapModel.AreaObj.SetLoadSurfacePressure("ALL", "DEAD", -1, .1, , , Group)



   'get area object surface pressure load

      ret = SapModel.AreaObj.GetLoadSurfacePressure("ALL", NumberItems, AreaName, LoadPat, Face, Value, PatternName, Group)



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

[SetLoadSurfacePressure](SetLoadSurfacePressure_%7BArea_Object%7D.md)

[DeleteLoadSurfacePressure](DeleteLoadSurfacePressure_%7BArea_Object%7D.md)

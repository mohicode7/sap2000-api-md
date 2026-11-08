# SetLoadSurfacePressure

## Syntax

SapObject.SapModel.AreaObj.SetLoadSurfacePressure

## VB6 Procedure

Function SetLoadSurfacePressure(ByVal Name As String, ByVal LoadPat As String, ByVal Face As Long, ByVal Value As Double, Optional ByVal PatternName As String = "", Optional ByVal Replace As Boolean = True, Optional ByVal ItemType As eItemType = Object) As Long

## Parameters

Name

The name of an existing area object or group, depending on the value of the ItemType item.

LoadPat

The name of a defined load pattern.

Face

This is -1, -2 or a nonzero, positive integer, indicating the area object face to which the specified load assignment applies.

-1 = Bottom face

-2 = Top face

>0 = Edge face

Note that edge face n is from area object point n to area object point n + 1. For example, edge face 2 is from area object point 2 to area object point 3.

Value

This is the surface pressure value. [F/L2]

PatternName

This is blank or the name of a defined joint pattern. If it is blank, the surface pressure load for the area object face is uniform over the face at the value specified by Value.

If PatternName is the name of a defined joint pattern, the surface pressure load for the area object face is based on the specified surface pressure value multiplied by the pattern value at the point objects that are part of the face.

Replace

If this item is True, all previous surface pressure loads, if any, assigned to the specified area object(s), in the specified load case, are deleted before making the new assignment.

ItemType

This is one of the following items in the eItemType enumeration:

Object = 0

Group = 1

SelectedObjects = 2

If this item is Object, the assignment is made to the area object specified by the Name item.

If this item is Group, the assignment is made to all area objects in the group specified by the Name item.

If this item is SelectedObjects, assignment is made to all selected area objects, and the Name item is ignored.

## Remarks

This function assigns surface pressure loads to area objects.

The function returns zero if the loads are successfully assigned; otherwise it returns a nonzero value.

## VBA Example



```
Sub AssignAreaObjectSurfacePressureLoad()

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



   'assign area object surface pressure load

      ret = SapModel.AreaObj.SetLoadSurfacePressure("ALL", "DEAD", -1, .1, , , Group)



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

[GetLoadSurfacePressure](GetLoadSurfacePressure_%7BArea_Object%7D.md)

[DeleteLoadSurfacePressure](DeleteLoadSurfacePressure_%7BArea_Object%7D.md)

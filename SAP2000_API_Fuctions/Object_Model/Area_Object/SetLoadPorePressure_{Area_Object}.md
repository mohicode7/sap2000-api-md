# SetLoadPorePressure

## Syntax

SapObject.SapModel.AreaObj.SetLoadPorePressure

## VB6 Procedure

Function SetLoadPorePressure(ByVal Name As String, ByVal LoadPat As String, ByVal Value As Double, Optional ByVal PatternName As String = "", Optional ByVal Replace As Boolean = True, Optional ByVal ItemType As eItemType = Object) As Long

## Parameters

Name

The name of an existing area object or group, depending on the value of the ItemType item.

LoadPat

The name of a defined load pattern.

Value

This is the pore pressure value. [F/L2]

PatternName

This is blank or the name of a defined joint pattern. If it is blank, the pore pressure load for the area object is uniform over the object at the value specified by Value.

If PatternName is the name of a defined joint pattern, the pore pressure load for the area object is based on the specified pore pressure value multiplied by the pattern value at the point objects that define the area object.

Replace

If this item is True, all previous pore pressure loads, if any, assigned to the specified area object(s), in the specified load case, are deleted before making the new assignment.

ItemType

This is one of the following items in the eItemType enumeration:

Object = 0

Group = 1

SelectedObjects = 2

If this item is Object, the assignment is made to the area object specified by the Name item.

If this item is Group, the assignment is made to all area objects in the group specified by the Name item.

If this item is SelectedObjects, assignment is made to all selected area objects, and the Name item is ignored.

## Remarks

This function assigns pore pressure loads to area objects.

The function returns zero if the loads are successfully assigned; otherwise it returns a nonzero value.

## VBA Example



```
Sub AssignAreaObjectPorePressureLoad()

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



   'assign area object pore pressure load

      ret = SapModel.AreaObj.SetLoadPorePressure("ALL", "DEAD", .1, , , Group)



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

[GetLoadPorePressure](GetLoadPorePressure_%7BArea_Object%7D.md)

[DeleteLoadPorePressure](DeleteLoadPorePressure_%7BArea_Object%7D.md)

# GetLoadWindPressure (Note: Newer function available)

## Syntax

SapObject.SapModel.AreaObj.GetLoadWindPressure

## VB6 Procedure

Function GetLoadWindPressure(ByVal Name As String, ByRef NumberItems As Long, ByRef AreaName() As String, ByRef LoadPat() As String, ByRef MyType() As Long, ByRef Cp() As Double, Optional ByVal ItemType As eItemType = Object) As Long

## Parameters

Name

The name of an existing area object or group, depending on the value of the ItemType item.

NumberItems

The total number of wind pressure loads retrieved for the specified area objects.

AreaName

This is an array that includes the name of the area object associated with each wind pressure load.

LoadPat

This is an array that includes the name of the load pattern associated with each wind pressure load.

MyType

This is an array that includes either 1 or 2, indicating the wind pressure type.

1 = Windward, pressure varies over height

2 = Other, pressure is constant over height

Cp

This is an array that includes the wind pressure coefficient value.

ItemType

This is one of the following items in the eItemType enumeration:

Object = 0

Group = 1

SelectedObjects = 2

If this item is Object, the assignments are retrieved for the area object specified by the Name item.

If this item is Group, the assignments are retrieved for all area objects in the group specified by the Name item.

If this item is SelectedObjects, assignments are retrieved for all selected area objects, and the Name item is ignored.

## Remarks

This function retrieves the wind pressure load assignments to area objects.

The function returns zero if the load assignments are successfully retrieved; otherwise it returns a nonzero value.

## VBA Example



```
Sub GetAreaObjectWindPressureLoad()

       'dimension variables

          Dim SapObject as cOAPI

          Dim SapModel As cSapModel

          Dim ret As Long

          Dim NumberItems As Long

          Dim AreaName() As String

          Dim LoadPat() As String

          Dim MyType() As Long

          Dim Cp() As Double



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



       'assign area object wind pressure load

          ret = SapModel.AreaObj.SetLoadWindPressure("ALL", "DEAD", 1, 0.8, Group)



       'get area object wind pressure load

          ret = SapModel.AreaObj.GetLoadWindPressure("ALL", NumberItems, AreaName, LoadPat, MyType, Cp, Group)



       'close Sap2000

          SapObject.ApplicationExit False

          Set SapModel = Nothing

          Set SapObject = Nothing

    End Sub
  
```



## Release Notes

Initial release in version 11.00.

Changed nomenclature from Load Cases, Analysis Cases and Response Combinations to Load Patterns, Load Cases and Load Combinations, respectively, in version 12.00.

The function is obsolete and has been superseded by [GetLoadWindPressure\_1](../Object_Model/Area_Object/GetLoadWindPressure_1.md) as of v22.1.0. This function is maintained for backward compatibility.

## See Also

[SetLoadWindPressure](../Object_Model/Area_Object/SetLoadWindPressure.md)

[DeleteLoadWindPressure](../Object_Model/Area_Object/DeleteLoadWindPressure.md)

# GetLoadWindPressure\_1

## Syntax

SapObject.SapModel.AreaObj.GetLoadWindPressure\_1

## VB6 Procedure

Function GetLoadWindPressure(ByVal Name As String, ByRef NumberItems As Long, ByRef AreaName() As String, ByRef LoadPat() As String, ByRef MyType() As Long, ByRef Cp() As Double, ByRef DistributionType() As Long, Optional ByVal ItemType As eItemType = Object) As Long

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

DistributionType

This is either 1 or 2, indicating the distribution type.

1 = To Joints  
2 = To Frames – One-way  
3 = To Frames – Two-way

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

      Dim DistributionType() As Long



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

          ret = SapModel.AreaObj.SetLoadWindPressure("ALL", "DEAD", 1, 0.8, 1, Group)



       'get area object wind pressure load

          ret = SapModel.AreaObj.GetLoadWindPressure("ALL", NumberItems, AreaName, LoadPat, MyType, Cp, DistributionType, Group)



       'close Sap2000

          SapObject.ApplicationExit False

          Set SapModel = Nothing

          Set SapObject = Nothing

    End Sub
  
```



## Release Notes

Initial release in version 22.1.0.

## See Also

[SetLoadWindPressure\_1](SetLoadWindPressure_1.md)

[DeleteLoadWindPressure](DeleteLoadWindPressure.md)

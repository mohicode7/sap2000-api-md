# SetLoadWindPressure\_1

## Syntax

SapObject.SapModel.AreaObj.SetLoadWindPressure\_1

## VB6 Procedure

Function SetLoadWindPressure(ByVal Name As String, ByVal
LoadPat As String, ByVal MyType As Long, ByVal Cp As Double, ByVal DistributionType
As Long, Optional ByVal ItemType As eItemType = Object) As Long

## Parameters

Name

The name of an existing area object or group, depending
on the value of the ItemType item.

LoadPat

The name of a defined load pattern.

MyType

This is either 1 or 2, indicating the wind pressure
type.

1 = Windward, pressure
varies over height

2 = Other, pressure is constant
over height

Cp

This is the wind pressure coefficient.

DistributionType

This is either 1 or 2, indicating the distribution type.

1 = To Joints  
2 = To Frames – One-way  
3 = To Frames – Two-way

**ItemType**

This is one of the following items in the eItemType
enumeration:

Object = 0

Group = 1

SelectedObjects = 2

If this item is Object, the assignment is made to the
area object specified by the Name item.

If this item is Group, the assignment is made to all
area objects in the group specified by the Name item.

If this item is SelectedObjects, assignment is made
to all selected area objects, and the Name item is ignored.

## Remarks

This function assigns wind pressure loads to area objects.

The function returns zero if the loads are successfully
assigned; otherwise it returns a nonzero value.

## VBA Example



```
Sub AssignAreaObjectWindPressureLoad()

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

      ret = SapModel.File.NewWall(2, 48, 
 2, 48)



   'assign area object wind pressure load

      ret = SapModel.AreaObj.SetLoadWindPressure("ALL", 
 "DEAD", 1, 0.8, 1, Group)



   'close Sap2000

      SapObject.ApplicationExit False

      Set SapModel = Nothing

      Set SapObject = Nothing

End Sub
```



## Release Notes

Initial release in version 22.1.0.

## See Also

[GetLoadWindPressure\_1](GetLoadWindPressure_1.md)

[DeleteLoadWindPressure](DeleteLoadWindPressure.md)

# GetLoadTemperature

## Syntax

SapObject.SapModel.AreaObj.GetLoadTemperature

## VB6 Procedure

Function GetLoadTemperature(ByVal Name As String, ByRef NumberItems As Long, ByRef AreaName() As String, ByRef LoadPat() As String, ByRef MyType() As Long, ByRef Value() As Double, ByRef PatternName() As String, Optional ByVal ItemType As eItemType = Object) As Long

## Parameters

Name

The name of an existing area object or group, depending on the value of the ItemType item.

NumberItems

The total number of temperature loads retrieved for the specified area objects.

AreaName

This is an array that includes the name of the area object associated with each temperature load.

LoadPat

This is an array that includes the name of the load pattern associated with each temperature load.

MyType

This is an array that includes either 1 or 3, indicating the type of temperature load.

1 = Temperature

3 = Temperature gradient along local 3 axis

Value

This is an array that includes the temperature load value. [T] for MyType= 1 and [T/L] for MyType= 3

PatternName

This is an array that includes the joint pattern name, if any, used to specify the temperature load.

ItemType

This is one of the following items in the eItemType enumeration:

Object = 0

Group = 1

SelectedObjects = 2

If this item is Object, the assignments are retrieved for the area object specified by the Name item.

If this item is Group, the assignments are retrieved for all area objects in the group specified by the Name item.

If this item is SelectedObjects, assignments are retrieved for all selected area objects, and the Name item is ignored.

## Remarks

This function retrieves the temperature load assignments to area objects.

The function returns zero if the load assignments are successfully retrieved; otherwise it returns a nonzero value.

## VBA Example



```
Sub GetAreaObjectTemperatureLoad()

   'dimension variables

      Dim SapObject as cOAPI

      Dim SapModel As cSapModel

      Dim ret As Long

      Dim NumberItems As Long

      Dim AreaName() As String

      Dim LoadPat() As String

      Dim MyType() As Long

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



   'assign area object temperature load

      ret = SapModel.AreaObj.SetLoadTemperature("All", "DEAD", 1, 50, , , Group)



   'get area object temperature load

      ret = SapModel.AreaObj.GetLoadTemperature("ALL", NumberItems, AreaName, LoadPat, MyType, Value, PatternName, Group)



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

[SetLoadTemperature](SetLoadTemperature_%7BArea_Object%7D.md)

[DeleteLoadTemperature](DeleteLoadTemperature_%7BArea_Object%7D.md)

# SetLoadTemperature

## Syntax

SapObject.SapModel.CableObj.SetLoadTemperature

## VB6 Procedure

Function SetLoadTemperature(ByVal Name As String, ByVal LoadPat As String, ByVal Val As Double, Optional ByVal PatternName As String = "", Optional ByVal Replace As Boolean = True, Optional ByVal ItemType As eItemType = Object) As Long

## Parameters

Name

The name of an existing cable object or group, depending on the value of the ItemType item.

LoadPat

The name of a defined load pattern.

Val

This is the temperature change value. [T]

PatternName

This is blank or the name of a defined joint pattern. If it is blank, the temperature load for the cable object is uniform along the object at the value specified by Val.

If PatternName is the name of a defined joint pattern, the temperature load for the cable object is based on the specified temperature value multiplied by the pattern value at the joints at each end of the cable object.

Replace

If this item is True, all previous temperature loads, if any, assigned to the specified cable object(s), in the specified load case, are deleted before making the new assignment.

ItemType

This is one of the following items in the eItemType enumeration:

Object = 0

Group = 1

SelectedObjects = 2

If this item is Object, the assignment is made to the cable object specified by the Name item.

If this item is Group, the assignment is made to all cable objects in the group specified by the Name item.

If this item is SelectedObjects, assignment is made to all selected cable objects, and the Name item is ignored.

## Remarks

This function assigns temperature loads to cable objects.

The function returns zero if the loads are successfully assigned, otherwise it returns a nonzero value.

## VBA Example



```
Sub AssignCableTemperatureLoad()

   'dimension variables

      Dim SapObject as cOAPI

      Dim SapModel As cSapModel

      Dim ret As Long

      Dim Name As String



   'create Sap2000 object

      Set SapObject = CreateObject("CSI.SAP2000.API.SapObject")



   'start Sap2000 application

      SapObject.ApplicationStart



   'create SapModel object

      Set SapModel = SapObject.SapModel



   'initialize model

      ret = SapModel.InitializeNewModel



   'create model from template

      ret = SapModel.File.New2DFrame(PortalFrame, 3, 124, 3, 200)



   'add cable object by points

      ret = SapModel.CableObj.AddByPoint("1", "6", Name)



   'set cable data

      ret = SapModel.CableObj.SetCableData(Name, 7, 1, 0, 0, 24)



   'assign cable temperature load

      ret = SapModel.CableObj.SetLoadTemperature("ALL", "DEAD", 50, , , Group)



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

[GetLoadTemperature](GetLoadTemperature_%7BCable_Object%7D.md)

[DeleteLoadTemperature](DeleteLoadTemperature_%7BCable_Object%7D.md)

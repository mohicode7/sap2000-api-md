# GetLoadGravity

## Syntax

SapObject.SapModel.CableObj.GetLoadGravity

## VB6 Procedure

Function GetLoadGravity(ByVal Name As String, ByRef NumberItems As Long, ByRef CableName() As String, ByRef LoadPat() As String, ByRef CSys() As String, ByRef x() As Double, ByRef y() As Double, ByRef z() As Double, Optional ByVal ItemType As eItemType = Object) As Long

## Parameters

Name

The name of an existing cable object or group, depending on the value of the ItemType item.

NumberItems

The total number of gravity loads retrieved for the specified cable objects.

CableName

This is an array that includes the name of the cable object associated with each gravity load.

LoadPat

This is an array that includes the name of the coordinate system in which the gravity load multipliers are specified.

CSys

This is an array that includes the name of the coordinate system associated with each gravity load.

x, y, z

These are arrays of gravity load multipliers in the x, y and z directions of the specified coordinate system.

ItemType

This is one of the following items in the eItemType enumeration:

Object = 0

Group = 1

SelectedObjects = 2

If this item is Object, the assignments are retrieved for the cable object specified by the Name item.

If this item is Group, the assignments are retrieved for all cable objects in the group specified by the Name item.

If this item is SelectedObjects, assignments are retrieved for all selected cable objects, and the Name item is ignored.

## Remarks

This function retrieves the gravity load assignments to cable objects.

The function returns zero if the load assignments are successfully retrieved, otherwise it returns a nonzero value.

## VBA Example



```
Sub GetCableGravityLoad()

   'dimension variables

      Dim SapObject as cOAPI

      Dim SapModel As cSapModel

      Dim ret As Long

      Dim NumberItems As Long

      Dim CableName() As String

      Dim LoadPat() As String

      Dim CSys() As String

      Dim x() As Double

      Dim y() As Double

      Dim z() As Double

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



   'assign cable gravity loads

      ret = SapModel.CableObj.SetLoadGravity("ALL", "DEAD", 0, 0, -1, , , Group)



   'get cable gravity load

      ret = SapModel.CableObj.GetLoadGravity(Name, NumberItems, CableName, LoadPat, CSys, x, y, z)



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

[SetLoadGravity](SetLoadGravity_%7BCable_Object%7D.md)

[DeleteLoadGravity](DeleteLoadGravity_%7BCable_Object%7D.md)

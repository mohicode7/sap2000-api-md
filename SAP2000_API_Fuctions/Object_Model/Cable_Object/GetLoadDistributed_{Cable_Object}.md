# GetLoadDistributed

## Syntax

SapObject.SapModel.CableObj.GetLoadDistributed

## VB6 Procedure

Function GetLoadDistributed(ByVal Name As String, ByRef NumberItems As Long, ByRef CableName() As String, ByRef LoadPat() As String, ByRef MyType() As Long, ByRef CSys() As String, ByRef Dir() As Long, ByRef ByRef Value() As Double, Optional ByVal ItemType As eItemType = Object) As Long

## Parameters

Name

The name of an existing cable object or group, depending on the value of the ItemType item.

NumberItems

The total number of distributed loads retrieved for the specified cable objects.

CableName

This is an array that includes the name of the cable object associated with each distributed load.

LoadPat

This is an array that includes the name of the coordinate system in which the distributed loads are specified.

MyType

This is an array that includes 1 or 2, indicating the type of distributed load.

1 = Force

2 = Moment

CSys

This is an array that includes the name of the coordinate system in which each distributed load is defined. It may be Local or the name of a defined coordinate system.

Dir

This is 1, 2, 3, 4, 5, 6 or 10, indicating the direction of the load.

1 = Local 1 axis (only applies when CSys is Local)

2 = Local 2 axis (only applies when CSys is Local)

3 = Local 3 axis (only applies when CSys is Local)

4 = X direction (does not apply when CSys is Local)

5 = Y direction (does not apply when CSys is Local)

6 = Z direction (does not apply when CSys is Local)

10 = Gravity direction (only applies when CSys is Global)

The positive gravity direction (see Dir = 10) is in the negative Global Z direction.

Value

This is the load value of the distributed load. The distributed load is applied over the full length of the cable. [F/L] when MyType is 1 and [FL/L] when MyType is 2

ItemType

This is one of the following items in the eItemType enumeration:

Object = 0

Group = 1

SelectedObjects = 2

If this item is Object, the assignments are retrieved for the cable object specified by the Name item.

If this item is Group, the assignments are retrieved for all cable objects in the group specified by the Name item.

If this item is SelectedObjects, assignments are retrieved for all selected cable objects, and the Name item is ignored.

## Remarks

This function retrieves the distributed load assignments to cable objects. The loads are uniformly distributed over the full length of cable objects.

The function returns zero if the load assignments are successfully retrieved, otherwise it returns a nonzero value.

## VBA Example



```
Sub GetCableDistributedLoad()

   'dimension variables

      Dim SapObject as cOAPI

      Dim SapModel As cSapModel

      Dim ret As Long

      Dim NumberItems As Long

      Dim CableName() As String

      Dim LoadPat() As String

      Dim MyType() As Long

      Dim CSys() As String

      Dim Dir() As Long

      Dim Value() As Double

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



   'assign cable distributed load

      ret = SapModel.CableObj.SetLoadDistributed(Name, "DEAD", 1, 10, 0.08)



   'get cable distributed loads

      ret = SapModel.CableObj.GetLoadDistributed("ALL", NumberItems, CableName, LoadPat, MyType, CSys, Dir, Value, Group)



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

[SetLoadDistributed](SetLoadDistributed_%7BCable_Object%7D.md)

[DeleteLoadDistributed](DeleteLoadDistributed_%7BCable_Object%7D.md)

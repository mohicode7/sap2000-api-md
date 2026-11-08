# SetLoadDistributed

## Syntax

SapObject.SapModel.CableObj.SetLoadDistributed

## VB6 Procedure

Function SetLoadDistributed(ByVal Name As String, ByVal LoadPat As String, ByVal MyType As Long, ByVal Dir As Long, ByVal Value As Double, Optional ByVal CSys As String = "Global", Optional ByVal Replace As Boolean = True, Optional ByVal ItemType As eItemType = Object) As Long

## Parameters

Name

The name of an existing cable object or group, depending on the value of the ItemType item.

LoadPat

The name of a defined load pattern.

MyType

This is 1 or 2, indicating the type of distributed load.

1 = Force per unit length

2 = Moment per unit length

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

CSys

This is Local or the name of a defined coordinate system. It is the coordinate system in which the loads are specified.

Replace

If this item is True, all previous loads, if any, assigned to the specified cable object(s), in the specified load pattern, are deleted before making the new assignment.

ItemType

This is one of the following items in the eItemType enumeration:

Object = 0

Group = 1

SelectedObjects = 2

If this item is Object, the assignment is made to the cable object specified by the Name item.

If this item is Group, the assignment is made to all cable objects in the group specified by the Name item.

If this item is SelectedObjects, assignment is made to all selected cable objects, and the Name item is ignored.

## Remarks

This function assigns uniform distributed loads over the full length of cable objects.

The function returns zero if the loads are successfully assigned, otherwise it returns a nonzero value.

## VBA Example



```
Sub AssignCableDistributedLoad()

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



   'assign cable distributed load

      ret = SapModel.CableObj.SetLoadDistributed(Name, "DEAD", 1, 10, 0.08)



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

[GetLoadDistributed](GetLoadDistributed_%7BCable_Object%7D.md)

[DeleteLoadDistributed](../../../DeleteLoadDistributedWithGUID.md)

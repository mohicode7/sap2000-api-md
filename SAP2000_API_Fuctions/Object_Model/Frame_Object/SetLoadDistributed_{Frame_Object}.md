# SetLoadDistributed

## Syntax

SapObject.SapModel.FrameObj.SetLoadDistributed

## VB6 Procedure

Function SetLoadDistributed(ByVal Name As String, ByVal LoadPat As String, ByVal MyType As Long, ByVal Dir As Long, ByVal Dist1 As Double, ByVal Dist2 As Double, ByVal Val1 As Double, ByVal Val2 As Double, Optional ByVal CSys As String = "Global", Optional ByVal RelDist As Boolean = True, Optional ByVal Replace As Boolean = True, Optional ByVal ItemType As eItemType = Object) As Long

## Parameters

Name

The name of an existing frame object or group, depending on the value of the ItemType item.

LoadPat

The name of a defined load pattern.

MyType

This is 1 or 2, indicating the type of distributed load.

1 = Force per unit length

2 = Moment per unit length

Dir

This is an integer between 1 and 11, indicating the direction of the load.

1 = Local 1 axis (only applies when CSys is Local)

2 = Local 2 axis (only applies when CSys is Local)

3 = Local 3 axis (only applies when CSys is Local)

4 = X direction (does not apply when CSys is Local)

5 = Y direction (does not apply when CSys is Local)

6 = Z direction (does not apply when CSys is Local)

7 = Projected X direction (does not apply when CSys is Local)

8 = Projected Y direction (does not apply when CSys is Local)

9 = Projected Z direction (does not apply when CSys is Local)

10 = Gravity direction (only applies when CSys is Global)

11 = Projected Gravity direction (only applies when CSys is Global)

The positive gravity direction (see Dir = 10 and 11) is in the negative Global Z direction.

Dist1

This is the distance from the I-End of the frame object to the start of the distributed load. This may be a relative distance (0 <= Dist1 <= 1) or an actual distance, depending on the value of the RelDist item. [L] when RelDist is False

Dist2

This is the distance from the I-End of the frame object to the end of the distributed load. This may be a relative distance (0 <= Dist2 <= 1) or an actual distance, depending on the value of the RelDist item. [L] when RelDist is False

Val1

This is the load value at the start of the distributed load. [F/L] when MyType is 1 and [FL/L] when MyType is 2

Val2

This is the load value at the end of the distributed load. [F/L] when MyType is 1 and [FL/L] when MyType is 2

CSys

This is Local or the name of a defined coordinate system. It is the coordinate system in which the loads are specified.

RelDist

If this item is True, the specified Dist item is a relative distance, otherwise it is an actual distance.

Replace

If this item is True, all previous distributed loads, if any, assigned to the specified frame object(s), in the specified load pattern, are deleted before making the new assignment.

ItemType

This is one of the following items in the eItemType enumeration:

Object = 0

Group = 1

SelectedObjects = 2

If this item is Object, the assignment is made to the frame object specified by the Name item.

If this item is Group, the assignment is made to all frame objects in the group specified by the Name item.

If this item is SelectedObjects, assignment is made to all selected frame objects, and the Name item is ignored.

## Remarks

This function assigns distributed loads to frame objects.

The function returns zero if the loads are successfully assigned, otherwise it returns a nonzero value.

## VBA Example



```
Sub AssignFrameDistributedLoad()

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

      ret = SapModel.File.New2DFrame(PortalFrame, 3, 124, 3, 200)



   'assign frame distributed loads

      ret = SapModel.FrameObj.SetLoadDistributed("15", "DEAD", 1, 10, 0, 1, 0.08, 0.08)



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

[GetLoadDistributed](GetLoadDistributed_%7BFrame_Object%7D.md)

[DeleteLoadDistributed](DeleteLoadDistributed_%7BFrame_Object%7D.md)

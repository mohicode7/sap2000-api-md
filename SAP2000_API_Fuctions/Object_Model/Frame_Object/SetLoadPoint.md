# SetLoadPoint

## Syntax

SapObject.SapModel.FrameObj.SetLoadPoint

## VB6 Procedure

Function SetLoadPoint(ByVal Name As String, ByVal LoadPat
As String, ByVal MyType As Long, ByVal Dir As Long, ByVal Dist As Double,
ByVal Val As Double, Optional ByVal CSys As String = "Global",
Optional ByVal RelDist As Boolean = True, Optional ByVal Replace As Boolean
= True, Optional ByVal ItemType As eItemType = Object) As Long

## Parameters

Name

The name of an existing frame object or group, depending
on the value of the ItemType item.

LoadPat

The name of a defined load pattern.

MyType

This is 1 or 2, indicating the type of point load.

1 = Force

2 = Moment

Dir

This is an integer between 1 and 11, indicating the
direction of the load.

1 = Local 1 axis (only
applies when CSys is Local)

2 = Local 2 axis (only
applies when CSys is Local)

3 = Local 3 axis (only
applies when CSys is Local)

4 = X direction (does
not apply when CSys is Local)

5 = Y direction (does
not apply when CSys is Local)

6 = Z direction (does
not apply when CSys is Local)

7 = Not Used - Projected
X direction (does not apply when CSys is Local)

8 = Not Used - Projected
Y direction (does not apply when CSys is Local)

9 = Not Used - Projected
Z direction (does not apply when CSys is Local)

10 = Gravity direction
(only applies when CSys is Global)

11 = Not Used - Projected
Gravity direction (only applies when CSys is Global)

The positive gravity direction (see Dir = 10 and 11)
is in the negative Global Z direction.

Dist

This is the distance from the I-End of the frame object
to the load location. This may be a relative distance (0 <= Dist <=
1) or an actual distance, depending on the value of the RelDist item.
[L] when RelDist is False

Val

This is the value of the point load. [F] when MyType
is 1 and [FL] when MyType is 2

CSys

This is Local or the name of a defined coordinate system.
It is the coordinate system in which the loads are specified.

RelDist

If this item is True, the specified Dist item is a relative
distance, otherwise it is an actual distance.

Replace

If this item is True, all previous loads, if any, assigned
to the specified frame object(s), in the specified load pattern, are deleted
before making the new assignment.

ItemType

This is one of the following items in the eItemType
enumeration:

Object = 0

Group = 1

SelectedObjects = 2

If this item is Object, the assignment is made to the
frame object specified by the Name item.

If this item is Group, the assignment is made to all
frame objects in the group specified by the Name item.

If this item is SelectedObjects, assignment is made
to all selected frame objects, and the Name item is ignored.

## Remarks

This function assigns point loads to frame objects.

The function returns zero if the loads are successfully
assigned, otherwise it returns a nonzero value.

## VBA Example



```
Sub AssignFramePointLoad()

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

      ret = SapModel.File.New2DFrame(PortalFrame, 
 3, 124, 3, 200)



   'assign frame point loads

      ret = SapModel.FrameObj.SetLoadPoint("15", 
 "DEAD", 1, 10, .5, 20)



   'close Sap2000

      SapObject.ApplicationExit False

      Set SapModel = Nothing

      Set SapObject = Nothing

End Sub
```



## Release Notes

Initial release in version 11.00.

Changed nomenclature from Load Cases, Analysis Cases
and Response Combinations to Load Patterns, Load Cases and Load Combinations,
respectively, in version 12.00.

Changed descriptions for Dir 7, 8, 9 and 11 to "Not
Used" in version 25.0.0

## See Also

[GetLoadPoint](GetLoadPoint_%7BFrame_Object%7D.md)

[DeleteLoadPoint](DeleteLoadPoint.md)

# GetLoadPoint

## Syntax

SapObject.SapModel.FrameObj.GetLoadPoint

## VB6 Procedure

Function GetLoadPoint(ByVal Name As String, ByRef NumberItems
As Long, ByRef FrameName() As String, ByRef LoadPat() As String, ByRef
MyType() As Long, ByRef CSys() As String, ByRef Dir() As Long, ByRef RelDist()
As Double, ByRef Dist() As Double, ByRef Val() As Double, Optional ByVal
ItemType As eItemType = Object) As Long

## Parameters

Name

The name of an existing frame object or group depending
on the value of the ItemType item.

NumberItems

The total number of point loads retrieved for the specified
frame objects.

FrameName

This is an array that includes the name of the frame
object associated with each point load.

LoadPat

This is an array that includes the name of the coordinate
system in which the point loads are specified.

MyType

This is an array that includes 1 or 2, indicating the
type of point load.

1 = Force

2 = Moment

CSys

This is an array that includes the name of the coordinate
system in which each point load is defined. It may be Local or the name
of a defined coordinate system.

Dir

This is an array that includes an integer between 1
and 11 indicating the direction of the load.

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

7 = Not used - Projected
X direction (does not apply when CSys is Local)

8 = Not used - Projected
Y direction (does not apply when CSys is Local)

9 = Not used - Projected
Z direction (does not apply when CSys is Local)

10 = Gravity direction
(only applies when CSys is Global)

11 = Not used - Projected
Gravity direction (only applies when CSys is Global)

The positive gravity direction (see Dir = 10 and 11)
is in the negative Global Z direction.

RelDist

This is an array that includes the relative distance
from the I-End of the frame object to the location where the point load
is applied.

Dist

This is an array that includes the actual distance from
the I-End of the frame object to the location where the point load is
applied. [L]

Val

This is an array that includes the value of the point
load. [F] when MyType is 1 and [FL] when MyType is 2

ItemType

This is one of the following items in the eItemType
enumeration:

Object = 0

Group = 1

SelectedObjects = 2

If this item is Object, the assignments are retrieved
for the frame object specified by the Name item.

If this item is Group, the assignments are retrieved
for all frame objects in the group specified by the Name item.

If this item is SelectedObjects, assignments are retrieved
for all selected frame objects, and the Name item is ignored.

## Remarks

This function retrieves the point load assignments to
frame objects.

The function returns zero if the load assignments are
successfully retrieved, otherwise it returns a nonzero value.

## VBA Example



```
Sub GetFramePointLoad()

   'dimension variables

      Dim SapObject as cOAPI

      Dim SapModel As cSapModel

      Dim ret As Long

      Dim NumberItems As Long

      Dim FrameName() As String

      Dim LoadPat() As String

      Dim MyType() As Long

      Dim CSys() As String

      Dim Dir() As Long

      Dim RelDist() As Double

      Dim Dist() As Double

      Dim Val() As Double



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

      ret = SapModel.FrameObj.SetLoadPoint("14", 
 "DEAD", 1, 10, .5, 20)

      ret = SapModel.FrameObj.SetLoadPoint("15", 
 "DEAD", 1, 10, .5, 20)



   'get frame point loads

      ret = SapModel.FrameObj.GetLoadPoint("ALL", 
 NumberItems, FrameName, LoadPat, MyType, CSys, Dir, RelDist, Dist, Val, 
 Group)



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

[SetLoadPoint](SetLoadPoint.md)

[DeleteLoadPoint](DeleteLoadPoint.md)

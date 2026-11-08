# SetLocalAxesAdvanced

## Syntax

SapObject.SapModel.FrameObj.SetLocalAxesAdvanced

## VB6 Procedure

Function SetLocalAxesAdvanced(ByVal Name As String, ByVal Active As Boolean, ByVal Plane2 As Long, ByVal PlVectOpt As Long, ByVal PlCSys As String, ByRef PlDir() As Long, ByRef PlPt() As String, ByRef PlVect() As Double, Optional ByVal ItemType As eItemType = Object) As Long

## Parameters

Name

The name of an existing frame object or group depending on the value of the ItemType item.

Active

This is True if advanced local axes exist.

Plane2

This is 12 or 13, indicating that the local plane determined by the plane reference vector is the 1-2 plane or the 1-3 plane. This item applies only when the Active item is True.

PlVectOpt

This is 1, 2, or 3, indicating the plane reference vector option. This item applies only when the Active item is True.

1 = Coordinate direction

2 = Two joints

3 = User vector

PlCSys

The coordinate system used to define the plane reference vector coordinate directions and the plane user vector. This item applies when the Active item is True and the PlVectOpt item is 1 or 3.

PlDir

This is an array dimensioned to 1 (2 integers), indicating the plane reference vector primary and secondary coordinate directions, PlDir(0) and PlDir(1) respectively, taken at the object center in the specified coordinate system and used to determine the plane reference vector. This item applies when the Active item is True and the PlVectOpt item is 1. Possible coordinate direction values are:

1 = +X      -1 = -X

2 = +Y      -2 = -Y

3 = +Z      -3 = -Z

4 = +CR     -4 = -CR

5 = +CA     -5 = -CA

6 = +CZ     -6 = -CZ

7 = +SR     -7 = -SR

8 = +SA     -8 = -SA

9 = +SB     -9 = -SB

PlPt

This is an array dimensioned to 1 (2 strings), indicating the labels of two joints that define the plane reference vector. Either of these joints may be specified as None to indicate the center of the specified object.  If both joints are specified as None, they are not used to define the plane reference vector. This item applies when the Active item is True and the PlVectOpt item is 2.

PlVect

This is an array dimensioned to 2 (3 doubles) that defines the plane reference vector. This item applies when the Active item is True and the PlVectOpt item is 3.

ItemType

This is one of the following items in the eItemType enumeration:

Object = 0

Group = 1

Selection = 2

If this item is Object, the assignment is made to the frame object specified by the Name item.

If this item is Group, the assignment is made to all frame objects in the group specified by the Name item.

If this item is Selection, assignment is made to all selected frame objects and the Name item is ignored.

## Remarks

This function assigns advanced local axes to frame objects.

The function returns zero if the advanced local axes assignments are assigned successfully; otherwise it returns a nonzero value.

## VBA Example



```
Sub AssignFrameAdvancedLocalAxes()

   'dimension variables

      Dim SapObject as cOAPI

      Dim SapModel As cSapModel

      Dim ret As Long

      Dim MyPlDir(1) As Long

      Dim MyPlPt(1) As String

      Dim MyPlVect(2) As Double



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



   'assign frame advanced local axes

      MyPlDir(0) = 2

      MyPlDir(1) = 3

      ret = SapModel.FrameObj.SetLocalAxesAdvanced("3", True, 12, 1, "Global", MyPlDir, MyPlPt, MyPlVect)



   'close Sap2000

      SapObject.ApplicationExit False

      Set SapModel = Nothing

      Set SapObject = Nothing

End Sub
```



## Release Notes

Initial release in version 14.00.

## See Also

[GetLocalAxesAdvanced](GetLocalAxesAdvanced_Frame_Object.md)

[GetLocalAxes](GetLocalAxes_%7BFrame_Object%7D.md)

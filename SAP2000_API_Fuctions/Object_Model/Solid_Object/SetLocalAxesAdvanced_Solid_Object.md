# SetLocalAxesAdvanced

## Syntax

SapObject.SapModel.SolidObj.SetLocalAxesAdvanced

## VB6 Procedure

Function SetLocalAxesAdvanced(ByVal Name As String, ByVal Active As Boolean, ByVal AxVectOpt As Long, ByVal AxCSys As String, ByRef AxDir() As Long, ByRef AxPt() As String, ByRef AxVect() As Double, ByVal Plane2 As Long, ByVal PlVectOpt As Long, ByVal PlCSys As String, ByRef PlDir() As Long, ByRef PlPt() As String, ByRef PlVect() As Double, Optional ByVal ItemType As eItemType = Object) As Long

## Parameters

Name

The name of an existing solid object or group, depending on the value of the ItemType item.

Active

This is True if advanced local axes exist.

AxVectOpt, PlVectOpt

This is 1, 2, or 3, indicating the axis/plane reference vector option. This item applies only when the Active item is True.

1 = Coordinate direction

2 = Two joints

3 = User vector

AxCSys, PlCSys

The coordinate system used to define the axis/plane reference vector coordinate directions and the axis/plane user vector. This item applies when the Active item is True and the AxVectOpt/PlVectOpt item is 1 or 3.

AxDir, PlDir

This is an array dimensioned to 1 (2 integers) indicating the axis/plane reference vector primary and secondary coordinate directions, PlDir(0) and PlDir(1) respectively, taken at the object center in the specified coordinate system and used to determine the plane reference vector. This item applies when the Active item is True and the AxVectOpt/PlVectOpt item is 1. Possible coordinate direction values are:

1 = +X        -1 = -X

2 = +Y        -2 = -Y

3 = +Z        -3 = -Z

4 = +CR     -4 = -CR

5 = +CA     -5 = -CA

6 = +CZ     -6 = -CZ

7 = +SR     -7 = -SR

8 = +SA     -8 = -SA

9 = +SB     -9 = -SB

AxPt, PlPt

This is an array dimensioned to 1 (2 strings) indicating the labels of two joints that define the axis/plane reference vector. Either of these joints may be specified as None to indicate the center of the specified object.  If both joints are specified as None, they are not used to define the plane reference vector. This item applies when the Active item is True and the AxVectOpt/PlVectOpt item is 2.

AxVect, PlVect

This is an array dimensioned to 2 (3 doubles) that defines the axis/plane reference vector. This item applies when the Active item is True and the AxVectOpt/PlVectOpt item is 3.

Plane2

This is 12, 13, 21, 23, 31 or 32, indicating that the local plane determined by the plane reference vector is the 1-2, 1-3, 2-1, 2-3, 3-1, or 3-2 plane. This item applies only when the Active item is True.

ItemType

This is one of the following items in the eItemType enumeration:

Object = 0

Group = 1

Selection = 2

If this item is Object, the assignment is made to the solid object specified by the Name item.

If this item is Group, the assignment is made to all solid objects in the group specified by the Name item.

If this item is Selection, assignment is made to all selected solid objects and the Name item is ignored.

## Remarks

This function assigns advanced local axes to solid objects.

The function returns zero if the advanced local axes assignments are assigned successfully; otherwise it returns a nonzero value.

## VBA Example



```
Sub AssignSolidAdvancedLocalAxes()

   'dimension variables

      Dim SapObject as cOAPI

      Dim SapModel As cSapModel

      Dim ret As Long

      Dim MyAxDir(1) As Long

      Dim MyAxPt(1) As String

      Dim MyAxVect(2) As Double

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

      ret = SapModel.File.NewSolidBlock(300, 400, 200, , , 2, 2, 2)



   'assign solid advanced local axes

      MyAxVect(0)=0.707

      MyAxVect(1)=0.707

      MyAxVect(2)=0

      MyPlDir(0) = 2

      MyPlDir(1) = 3

      ret = SapModel.SolidObj.SetLocalAxesAdvanced("ALL", True, 3, "Global", MyAxDir, MyAxPt, MyAxVect, 12, 1, "Global", MyPlDir, MyPlPt, MyPlVect, Group)



   'close Sap2000

      SapObject.ApplicationExit False

      Set SapModel = Nothing

      Set SapObject = Nothing

End Sub
```



## Release Notes

Initial release in version 14.00.

## See Also

[GetLocalAxesAdvanced](GetLocalAxesAdvanced_Solid_Object.md)

[GetLocalAxes](GetLocalAxes_%7BSolid_Object%7D.md)

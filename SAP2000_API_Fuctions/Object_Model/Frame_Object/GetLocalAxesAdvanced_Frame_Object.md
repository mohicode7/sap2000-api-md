# GetLocalAxesAdvanced

## Syntax

SapObject.SapModel.FrameObj.GetLocalAxesAdvanced

## VB6 Procedure

Function GetLocalAxesAdvanced(ByVal Name As String, ByRef Active As Boolean, ByRef Plane2 As Long, ByRef PlVectOpt As Long, ByRef PlCSys As String, ByRef PlDir() As Long, ByRef PlPt() As String, ByRef PlVect() As Double) As Long

## Parameters

Name

The name of an existing frame object.

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

## Remarks

This function retrieves the advanced local axes assignments to frame objects.

The function returns zero if the advanced local axes assignments are successfully retrieved; otherwise it returns a nonzero value.

## VBA Example



```
Sub GetFrameAdvancedLocalAxes()

   'dimension variables

      Dim SapObject as cOAPI

      Dim SapModel As cSapModel

      Dim ret As Long

      Dim MyPlDir(1) As Long

      Dim MyPlPt(1) As String

      Dim MyPlVect(2) As Double

      Dim Ang As Double

      Dim Advanced As Boolean

      Dim Active As Boolean

      Dim Plane2 As Long

      Dim PlVectOpt As Long

      Dim PlCSys As String

      Dim PlDir() As Long

      Dim PlPt() As String

      Dim PlVect() As Double



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



   'get frame local axis angle

      ret = SapModel.FrameObj.GetLocalAxes("3", Ang, Advanced)



   'get frame advanced local axes data

      If Advanced Then

         ret = SapModel.FrameObj.GetLocalAxesAdvanced("3", Active, Plane2, PlVectOpt, PlCSys, PlDir, PlPt, PlVect)

      End If



   'close Sap2000

      SapObject.ApplicationExit False

      Set SapModel = Nothing

      Set SapObject = Nothing

End Sub
```



## Release Notes

Initial release in version 14.00.

## See Also

[SetLocalAxesAdvanced](SetLocalAxesAdvanced_Frame_Object.md)

[SetLocalAxes](SetLocalAxes_%7BFrame_Object%7D.md)

[GetLocalAxes](GetLocalAxes_%7BFrame_Object%7D.md)

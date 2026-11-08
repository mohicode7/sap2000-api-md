# SetCurved

## Syntax

SapObject.SapModel.FrameObj.SetCurved

## VB6 Procedure

Function SetCurved(ByVal Name As String, ByVal MyType As Long, ByVal x As Double, ByVal y As Double, ByVal z As Double, ByVal PointName As String, ByVal Radius As Double, ByVal NumSegs As Long, Optional ByVal CSys As String = "Global") As Long

## Parameters

Name

The name of a defined curved frame object.

MyType

This is 1, 2, 3, 4, or 5, indicating the curved frame type.

1 = Circular Arc Specified by a Third Point Name

2 = Circular Arc Specified by Third Point Coordinates

3 = Circular Arc Specified by Planar Point Coordinates and Radius

4 = Parabolic Arc Specified by a Third Point Name

5 = Parabolic Arc Specified by Third Point Coordinates

MyTypes 1, 2, 4, and 5 all define the curve by three points. The three points are the two end point of the frame object and a third point defined by naming an existing point object or specifying point coordinates.

MyType 3 defines a circular curved frame by it end points, the coordinates of another point that lies in the plane of the curve but not necessarily on the curved frame, and a curve radius.

x, y, z

These are point coordinates in the coordinate system specified by CSys. [L]

For MyType 1 and 4 these items do not apply.

For MyType 2 and 5 these are the coordinates of the third point on the curved frame.

For MyType 3 these are the coordinates of the planar point that lies in the plane of the curved frame.

PointName

This is the name of the point object that is the third point on the curved frame. This item applies for MyType 1 and 4. It does not apply for MyType 2, 3 and 5.

Radius

The radius of the circular curved frame. This item only applies for MyType 3. [L]

NumSegs

This is the number of segments into which the program internally divides the curved frame.

CSys

This is the coordinate system in which the coordinates x, y and z are defined.

## Remarks

This function changes the curve data for a curved frame object and sets straight frame objects to be curved.

The function returns zero if the frame object type is successfully set, otherwise it returns a nonzero value.

## VBA Example



```
Sub SetFrameCurved()

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



   'set frame curved

      ret = SapModel.FrameObj.SetCurved("13", 1, 0, 0, 0, "1", 0, 16)

      ret = SapModel.FrameObj.SetCurved("14", 2, -200, 0, 176, "", 0, 16)

      ret = SapModel.FrameObj.SetCurved("15", 3, 0, 0, 0, "", 100, 16)

      ret = SapModel.FrameObj.SetCurved("16", 4, 0, 0, 0, "3", 0, 16)

      ret = SapModel.FrameObj.SetCurved("17", 5, 0, 0, 176, "", 0, 16)



   'close Sap2000

      SapObject.ApplicationExit False

      Set SapModel = Nothing

      Set SapObject = Nothing

End Sub
```



## Release Notes

Initial release in version 11.00.

Modified optional argument CSys to be ByVal in version 12.0.1.

## See Also

[SetStraight](SetStraight.md)

[GetCurved](../../Obsolete_Functions/GetCurved.md)

[GetType](GetType_%7BFrame_Object%7D.md)

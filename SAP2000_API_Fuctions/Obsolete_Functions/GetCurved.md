# GetCurved

## Syntax

SapObject.SapModel.FrameObj.GetCurved

## VB6 Procedure

Function GetCurved(ByRef NumberItems As Long, ByRef MyType() As Long, ByRef gx() As Double, ByRef gy() As Double, ByRef gz() As Double, ByRef PointName() As String, ByRef Radius() As Double, ByRef NumSegs() As Long) As Long

## Parameters

NumberItems

The number of curved frame objects returned.

MyType

This is an array that includes a numeric value indicating the curved frame type. The type is 1, 2, 3, 4, or 5.

1 = Circular Arc Specified by a Third Point Name

2 = Circular Arc Specified by Third Point Coordinates

3 = Circular Arc Specified by Planar Point Coordinates and Radius

4 = Parabolic Arc Specified by a Third Point Name

5 = Parabolic Arc Specified by Third Point Coordinates

MyTypes 1, 2, 4, and 5 all define the curve by three points. The three points are the two end point of the frame object and a third point defined by naming an existing point object or specifying point coordinates.

MyType 3 defines a circular curved frame by it end points, the coordinates of another point that lies in the plane of the curve but not necessarily on the curved frame, and a curve radius.

gx, gy, gz

These are arrays that include the point coordinates in the global coordinate system. [L]

For MyType 1 and 4 these items do not apply.

For MyType 2 and 5 these are the coordinates of the third point on the curved frame.

For MyType 3 these are the coordinates of the planar point that lies in the plane of the curved frame.

PointName

This is an array that includes the name of the point object that is the third point on the curved frame. This item applies for MyType 1 and 4. It does not apply for MyType 2, 3 and 5.

Radius

This is an array of the radii of the circular curved frame. This item only applies for MyType 3. [L]

NumSegs

This is an array that includes the number of segments into which the program internally divides the curved frame.

## Remarks

This function retrieves definition data for all curved frame objects and returns the data in arrays.

The function returns zero if the curved frame object data is successfully retrieved, otherwise it returns a nonzero value.

## VBA Example



```
Sub GetCurvedFrames()

   'dimension variables

      Dim SapObject as cOAPI

      Dim SapModel As cSapModel

      Dim ret As Long

      Dim NumberItems As Long

      Dim MyType() As Long

      Dim gx() As Double

      Dim gy() As Double

      Dim gz() As Double

      Dim PointName() As String

      Dim Radius() As Double

      Dim NumSegs() As Long



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



   'set frames curved

      ret = SapModel.FrameObj.SetCurved("13", 1, 0, 0, 0, "1", 0, 16)

      ret = SapModel.FrameObj.SetCurved("14", 2, -200, 0, 176, "", 0, 16)

      ret = SapModel.FrameObj.SetCurved("15", 3, 0, 0, 0, "", 100, 16)

      ret = SapModel.FrameObj.SetCurved("16", 4, 0, 0, 0, "3", 0, 16)

      ret = SapModel.FrameObj.SetCurved("17", 5, 0, 0, 176, "", 0, 16)



   'get curved frame data

      ret = SapModel.FrameObj.GetCurved(NumberItems, MyType, gx, gy, gz, PointName, Radius, NumSegs)



   'close Sap2000

      SapObject.ApplicationExit False

      Set SapModel = Nothing

      Set SapObject = Nothing

End Sub
```



## Release Notes

Initial release in version 11.00.

## See Also

[SetStraight](../Object_Model/Frame_Object/SetStraight.md)

[SetCurved](../Object_Model/Frame_Object/SetCurved.md)

[GetType](../Object_Model/Frame_Object/GetType_%7BFrame_Object%7D.md)

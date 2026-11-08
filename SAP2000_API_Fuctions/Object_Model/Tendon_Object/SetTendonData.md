# SetTendonData

## Syntax

SapObject.SapModel.TendonObj.SetTendonData

## VB6 Procedure

Function SetTendonData(ByVal Name As String, ByVal NumberPoints As Long, ByRef MyType() As Long, ByRef x() As Double, ByRef y() As Double, ByRef z() As Double, Optional ByVal CSys As String = "Global") As Long

## Parameters

Name

The name of a defined tendon object.

NumberPoints

The number of items used to define the tendon geometry.

MyType

This is an array of values that are 1, 3, 6, 7, 8, or 9, indicating the tendon geometry definition parameter for the specified point.

1 = Start of tendon

2 = The segment preceding the point is linear

6 = The specified point is the end of a parabola

7 = The specified point is an intermediate point on a parabola

8 = The specified point is the end of a circle

9 = The specified point is an intermediate point on a parabola

The first point should always have a MyType value of 1. If it is not equal to 1, the program uses 1 anyway.

MyType of 6 through 9 is based on using three points to calculate a parabolic or circular arc. MyType 6 and 8 use the specified point and the two previous points as the three points. MyType 7 and 9 use the specified point and the points just before and after the specified point as the three points.

x

This is an array of the X (or local 1) coordinate of each point in the coordinate system specified by CSys. [L]

y

This is an array of the Y (or local 2) coordinate of each point in the coordinate system specified by CSys. [L]

z

This is an array of the Z (or local 3) coordinate of each point in the coordinate system specified by CSys. [L]

CSys

This is the coordinate system in which the x, y and z coordinate parameters are defined. It is Local or the name of a defined coordinate system.

Local means that the point coordinates are in the local system of the specified tendon object with the origin assumed to be at the I-End of the tendon.

## Remarks

This function assigns the tendon geometric definition parameters to a tendon object.

The function returns zero if the tendon object is successfully defined, otherwise it returns a nonzero value. If the tendon object is not successfully defined, it may be deleted.

## VBA Example



```
Sub SetTendonGeometryData()

   'dimension variables

      Dim SapObject as cOAPI

      Dim SapModel As cSapModel

      Dim ret As Long

      Dim Name As String

      Dim NumberPoints As Long

      Dim MyType() As Long

      Dim x() As Double

      Dim y() As Double

      Dim z() As Double



   'create Sap2000 object

      Set SapObject = CreateObject("CSI.SAP2000.API.SapObject")



   'start Sap2000 application

      SapObject.ApplicationStart



   'create SapModel object

      Set SapModel = SapObject.SapModel



   'initialize model

      ret = SapModel.InitializeNewModel



   'create model from template

      ret = SapModel.File.New2DFrame(PortalFrame, 2, 144, 2, 288)



   'add tendon object by coordinates

      ret = SapModel.TendonObj.AddByCoord(-288, 0, 288, 288, 0, 288, Name)



   'set tendon geometry

      NumberPoints = 3

      ReDim MyType(NumberPoints - 1)

      ReDim x(NumberPoints - 1)

      ReDim y(NumberPoints - 1)

      ReDim z(NumberPoints - 1)

      MyType(0) = 1

      MyType(1) = 7

      MyType(2) = 6

      x(0) = 0:  y(0) = 0

      x(1) = 288:  y(1) = -12

      x(2) = 576:  y(2) = 0

      ret = SapModel.TendonObj.SetTendonData(Name, NumberPoints, MyType, x, y, z, "Local")



   'update view

      ret = SapModel.View.RefreshView(0, False)



   'close Sap2000

      SapObject.ApplicationExit False

      Set SapModel = Nothing

      Set SapObject = Nothing

End Sub
```



## Release Notes

Initial release in version 11.00.

## See Also

[AddByCoord](AddByCoord_%7BTendon_Object%7D.md)

[AddByPoint](AddByPoint_%7BTendon_Object%7D.md)

[GetTendonData](GetTendonData.md)

[GetTendonGeometry](GetTendonGeometry.md)

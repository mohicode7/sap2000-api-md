# GetLineElevPoints

## Syntax

SapObject.SapModel.GenRefLine.GetLineElevPoints

## VB6 Procedure

Function GetLineElevPoints(ByVal Name As String, ByRef NumberPoints As Long, ByRef CurveType() As Long, ByRef Value1() As Double, ByRef Value2() As Double, ByRef Value3() As Double, ByRef s() As Double, ByRef z() As Double, Optional ByVal CSys As String = "Global") As Long

## Parameters

Name

The name of an existing general reference line.

NumberPoints

The number of points used to define the general reference line elevation layout.

CurveType

This is an array of values indicating the general reference line elevation layout curve type for each point.

0 = None  
1 = Circular Curve  
2 = Highway Curve  
3 = Parabolic Curve  
4 = Bezier Curve  
5 = BSpline Curve  
6 = Bezier Curve Child Point  
7 = BSpline Curve Child Point

Value1

This is the value of a parameter used to define the general reference line layout. The item that Value1 represents depends on the CurveType item.

CurveType = 0: Not Used  
CurveType = 1: Curve Radius [L]  
CurveType = 2: Curve Radius [L]  
CurveType = 3: Angle measured from the horizontal, up station axis, to the axis of symmetry of the parabolic curve. [deg]  
CurveType = 4: Number of control points. This is currently hard-wired internally to 4.  
CurveType = 5: Number of control points.  
CurveType = 6: Not Used  
CurveType = 7: Not Used

Value2  
This is the value of a parameter used to define the general reference line layout. The item that Value2 represents depends on the CurveType item.

CurveType = 0: Not Used  
CurveType = 1: Not Used  
CurveType = 2: Curve length, including length of spirals on either end. [L]  
CurveType = 3: Rate at which the slope of a parabolic curve is changing in percent. [1/L]  
CurveType = 4: Number of discretization points.  
CurveType = 5: Number of discretization points.  
CurveType = 6: Not Used  
CurveType = 7: Not Used

Value3

This is the value of a parameter used to define the general reference line layout. The item that Value3 represents depends on the CurveType item.

CurveType = 0: Not Used  
CurveType = 1: Not Used  
CurveType = 2: Not Used  
CurveType = 3: Not Used  
CurveType = 4: Not Used  
CurveType = 5: Curve order.  
CurveType = 6: Not Used  
CurveType = 7: Not Used

s

This is an array of the station coordinate of each point in the coordinate system specified for the general reference line. [L]

z

This is an array of the Z coordinate of each point in the coordinate system specified for the general reference line. [L]

## Remarks

This function retrieves the general reference line elevation points and the associated parameters.

The function returns zero if the general reference line parameters are successfully retrieved; otherwise it returns a nonzero value.

## VBA Example



```
Sub GetElevPoints()

   'dimension variables

      Dim SapObject as cOAPI

      Dim SapModel As cSapModel

      Dim ret As Long

      Dim Name As String

      Dim NumberPoints As Long

      Dim CurveType() As Long

      Dim Value1() As Double

      Dim Value2() As Double

      Dim Value3() As Double

      Dim s() As Double

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



   'add new general reference line

      Name = "GRef1"

      ret = SapModel.GenRefLine.SetLine(Name, 120, 5)



   'set general reference line elev points data

      NumberPoints = 3

      ReDim CurveType(NumberPoints – 1)

      ReDim x(NumberPoints – 1)

      Redim y(NumberPoints – 1)

      CurveType(0) = 0

      CurveType(1) = 1

      CurveType(2) = 0

      s(0) = 0: z(0) = 0

      s(1) = 100: z(1) = 100

      s(2) = 200: z(2) = 100

      Value1(0) = 0: Value1(1) = 100: Value1(2) = 0

      Value2(0) = 0: Value2(1) = 0: Value2(2) = 0

      Value3(0) = 0: Value3(1) = 0: Value3(2) = 0

      ret = SapModel.GenRefLine.SetLineElevPoints(Name, NumberPoints, CurveType, Value1, Value2, Value3, s, z)



   'get general reference line plan points data

      ReDim CurveType(0)

      ReDim Value1(0)

      ReDim Value2(0)

      ReDim Value3(0)

      ReDim s(0)

      ReDim z(0)

      Ret = SapModel.GenRefLine.GetLinePlanPoints(Name, NumberPoints, Value1, Value2, Value3, s, z)  



   'close Sap2000

      SapObject.ApplicationExit False

      Set SapModel = Nothing

      Set SapObject = Nothing

End Sub
```



## Release Notes

Initial release in version 15.0.0.

## See Also

[SetLineElevPoints](SetLineElevPoints.md) [GetLine](GetLine_%7BGen_Ref_Line%7D.md) [GetLinePlanPoints](GetLinePlanPoints.md)

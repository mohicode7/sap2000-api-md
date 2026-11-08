# ExtrudeFrameToAreaRadial

## Syntax

SapObject.SapModel.EditGeneral.ExtrudeFrameToAreaRadial

## VB6 Procedure

Function ExtrudeFrameToAreaRadial(ByVal Name As String,
ByVal PropName As String, ByVal RotateAxis As Long, ByVal x As Double,
ByVal y As Double, ByVal z As Double, ByVal IncrementAng As Double, ByVal
TotalRise As Double, ByVal NumberAreas As Long, ByRef AreaName() As String,
Optional ByVal Remove As Boolean = True) As Long

## Parameters

Name

The name of an existing line object to be extruded.

PropName

This is Default, None or the name of a defined area
property to be used for the new extruded area objects.

RotateAxis

This is 0, 1 or 2, indicating the axis that the radial
extrusion is around.

0 = X axis

1 = Y axis

2 = Z axis

x, y, z

These are the x, y and z coordinates, in the present
coordinate system, of the point that the radial extrusion is around. For
rotation about the X axis the value of the x coordinate is irrelevant.
Similarly, for rotation about the Y and Z axes the y and z coordinates,
respectively, are irrelevant. [L]

IncrementAng

The angle is rotated by this amount for each added area
object. [deg]

TotalRise

The total rise over the full length of the extrusion.
[L]

NumberAreas

The number of area objects created when the specified
line object is extruded.

AreaName

This is an array of the name of each area object created
when the specified line object is extruded.

Remove

If this item is True, the straight frame object indicated
by the Name item is deleted after the extrusion is complete.

## Remarks

This function creates new area objects by radially extruding
a specified straight frame object into area objects.

The function returns zero if the extrusion is successful;
otherwise it returns a nonzero value.

## VBA Example



```
Sub RadialFrameExtrusionToAreas()

   'dimension variables

      Dim SapObject as cOAPI

      Dim SapModel As cSapModel

      Dim ret As Long

      Dim AreaName() As String



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
 2, 144, 3, 288)



   'radially extrude frame to areas

      ret = SapModel.EditGeneral.ExtrudeFrameToAreaRadial("10", 
 "Default", 2, 0, 0, 288, 30, 0, 6, AreaName)



   'refresh view

      ret = SapModel.View.RefreshView(0, 
 False)



   'close Sap2000

      SapObject.ApplicationExit False

      Set SapModel = Nothing

      Set SapObject = Nothing

End Sub
```



## Release Notes

Initial release in version 11.00.

Modified optional argument Remove to be ByVal in version
12.0.1.

## See Also

[ExtrudeFrameToAreaLinear](ExtrudeFrameToAreaLinear.md)

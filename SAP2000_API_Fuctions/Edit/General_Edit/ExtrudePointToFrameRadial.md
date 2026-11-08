# ExtrudePointToFrameRadial

## Syntax

SapObject.SapModel.EditGeneral.ExtrudePointToFrameRadial

## VB6 Procedure

Function ExtrudePointToFrameRadial(ByVal Name As String,
ByVal PropName As String, ByVal RotateAxis As Long, ByVal x As Double,
ByVal y As Double, ByVal z As Double, ByVal IncrementAng As Double, ByVal
TotalRise As Double, ByVal NumberFrames As Long, ByRef FrameName() As
String) As Long

## Parameters

Name

The name of an existing point object to be extruded.

PropName

This is Default, None or the name of a defined frame
section property to be used for the new extruded frame objects.

RotateAxis

This is 0, 1 or 2, indicating the axis that the radial
extrusion is around.

0 = X axis

1 = Y axis

2 = Z axis

x, y, z

These are the x, y and z coordinates, in the present
coordinate system, of the point that the radial extrusion is around. For
rotation about the X axis, the value of the x coordinate is irrelevant.
Similarly, for rotation about the Y and Z axes, the y and z coordinates,
respectively, are irrelevant. [L]

IncrementAng

The angle is rotated by this amount for each added frame
object. [deg]

TotalRise

The total rise over the full length of the extrusion.
[L]

NumberFrames

The number of frame objects created when the specified
point object is extruded.

FrameName

This is an array of the name of each frame object created
when the specified point object is extruded.

## Remarks

This function creates new frame objects by radially
extruding a specified point object into frame objects.

The function returns zero if the extrusion is successful;
otherwise it returns a nonzero value.

## VBA Example



```
Sub RadialPointExtrusionToFrames()

   'dimension variables

      Dim SapObject as cOAPI

      Dim SapModel As cSapModel

      Dim ret As Long

      Dim FrameName() As String





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
 2, 144, 2, 288)



   'radially extrude point to frames

      ret = SapModel.EditGeneral.ExtrudePointToFrameRadial("3", 
 "FSec1", 2, 0, 0, 288, 30, 0, 6, FrameName)



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

## See Also

[ExtrudePointToFrameLinear](ExtrudePointToFrameLinear.md)

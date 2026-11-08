# ExtrudeAreaToSolidRadial

## Syntax

SapObject.SapModel.EditGeneral.ExtrudeAreaToSolidRadial

## VB6 Procedure

Function ExtrudeAreaToSolidRadial(ByVal Name As String,
ByVal PropName As String, ByVal RotateAxis As Long, ByVal x As Double,
ByVal y As Double, ByVal z As Double, ByVal IncrementAng As Double, ByVal
TotalRise As Double, ByVal Number As Long, ByRef NumberSolids As Long,
ByRef SolidName() As String, Optional ByVal Remove As Boolean = True)
As Long

## Parameters

Name

The name of an existing area object to be extruded.

PropName

This is either Default or the name of a defined solid
property to be used for the new extruded solid objects.

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

The angle is rotated by this amount for each added solid
object. [deg]

TotalRise

The total rise over the full length of the extrusion.
[L]

Number

The number of angle increments for the extrusion.

NumberSolids

The number of solid objects created when the specified
area object is extruded. Usually this item is returned the same as the
Number item. However, in some cases, such as when an area object with
more than four sides is extruded, this item will be larger than the Number
item.

SolidName

This is an array of the name of each solid object created
when the specified area object is extruded.

Remove

If this item is True, the area object indicated by the
Name item is deleted after the extrusion is complete.

## Remarks

This function creates new solid objects by radially
extruding a specified area object into solid objects.

The function returns zero if the extrusion is successful;
otherwise it returns a nonzero value.

## VBA Example



```
Sub RadialAreaExtrusionToSolids()

   'dimension variables

      Dim SapObject as cOAPI

      Dim SapModel As cSapModel

      Dim ret As Long

      Dim NumberSolids As Long

      Dim SolidName() As String



   'create Sap2000 object

      Set SapObject = CreateObject("CSI.SAP2000.API.SapObject")



   'start Sap2000 application

      SapObject.ApplicationStart



   'create SapModel object

      Set SapModel = SapObject.SapModel



   'initialize model

      ret = SapModel.InitializeNewModel



   'create model from template

      ret = SapModel.File.NewWall(5, 48, 
 2, 48)



   'radially extrude area to solid

      ret = SapModel.EditGeneral.ExtrudeAreaToSolidRadial("2", 
 "Default", 2, 0, 0, 96, 30, 0, 6, NumberSolids, SolidName)



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

[ExtrudeAreaToSolidLinearNormal](ExtrudeAreaToSolidLinearNormal.md)

[ExtrudeAreaToSolidLinearUser](ExtrudeAreaToSolidLinearUser.md)

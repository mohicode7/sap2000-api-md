# GetSolidSegment

## Syntax

SapObject.SapModel.PropFrame.SDShape.GetSolidSegment

## VB6 Procedure

Function GetSolidSegment(ByVal Name As String, ByValShapeName As String, ByRefMatProp As String, ByRef Color As Long, ByRefXCenter As Double, ByRefYCenter As Double, ByRef Angle As Double, ByRefRotation As Double, ByRef Radius As Double) As Long

## Parameters

Name

The name of an existing frame section property that is a section designer section.

ShapeName

The name of an existing solid segment shape in the specified frame section property.

MatProp

The name of the material property for the shape.

Color

The fill color assigned to the shape.

XCenter

The X-coordinate of the center of the shape in the section designer coordinate system. [L]

YCenter

The Y-coordinate of the center of the shape in the section designer coordinate system. [L]

Angle

The angle between lines drawn from the center of the circle to the end points of the chord tat defines the segment. [deg]

Rotation

The counter clockwise rotation of the shape from its default orientation. [deg]

Radius

The radius of the circle defining the segment.

## Remarks

This function retrieves property data for a solid segment shape in a section designer section.

The function returns zero if the property data is successfully retrieved; otherwise it returns a nonzero value.

## VBA Example



```
Sub GetFrameSDPropSolidSegment()

   'dimension variables

      Dim SapObject as cOAPI

      Dim SapModel As cSapModel

      Dim ret As Long

      Dim Name As String

      Dim MatProp As String

      Dim Color As Long

      Dim XCenter As Double

      Dim YCenter As Double

      Dim Angle As Double

      Dim Rotation As Double

      Dim Radius As Double



   'create Sap2000 object

      Set SapObject = CreateObject("CSI.SAP2000.API.SapObject")



   'start Sap2000 application

      SapObject.ApplicationStart



   'createSapModelobject

      Set SapModel = SapObject.SapModel



   'initialize model

      ret = SapModel.InitializeNewModel



   'create model from template

      ret = SapModel.File.New2DFrame(PortalFrame, 2, 144, 2, 288)



   'add new section designer frame section property

      ret = SapModel.PropFrame.SetSDSection("SD1", "4000Psi")



   'add solid segment shape to new property

      ret = SapModel.PropFrame.SDShape.SetSolidSegment("SD1", "SH1", "4000Psi", 0, 0, 95, 0, 10, -1)



   'get solid segment shape property data

      ret = SapModel.PropFrame.SDShape.GetSolidSegment("SD1", "SH1", MatProp, Color, Xcenter, Ycenter, Angle, Rotation, Radius)



   'close Sap2000

      SapObject.ApplicationExit False

      Set SapModel = Nothing

      Set SapObject = Nothing

End Sub
```



## Release Notes

Initial release in version 12.00.

## See Also

[SetSolidSegment](SetSolidSegment.md)

# GetPlate

## Syntax

SapObject.SapModel.PropFrame.SDShape.GetPlate

## VB6 Procedure

Function GetPlate(ByVal Name As String, ByVal ShapeName As String, ByRef MatProp As String, ByRef Color As Long, ByRef XCenter As Double, ByRef YCenter As Double, ByRef thickness As Double, ByRef w As Double, ByRef Rotation As Double) As Long

## Parameters

Name

The name of an existing frame section property that is a section designer section.

ShapeName

The name of an existing Plate shape in the specified frame section property.

MatProp

The name of the material property for the shape.

Color

The fill color assigned to the shape.

XCenter

The X-coordinate of the center of the shape in the section designer coordinate system. [L]

YCenter

The Y-coordinate of the center of the shape in the section designer coordinate system. [L]

thickness

The thickness of the plate. [L]

w

The width of the Plate. [L]

Rotation

The counter clockwise rotation of the shape from its default orientation. [deg]

## Remarks

This function retrieves property data for an Plate shape in a section designer section.

The function returns zero if the property data is successfully retrieved; otherwise it returns a nonzero value.

## VBA Example



```
Sub GetFrameSDPropPlate()

   'dimension variables

      Dim SapObject as cOAPI

      Dim SapModel As cSapModel

      Dim ret As Long

      Dim MatProp As String

      Dim Color As Long

      Dim XCenter As Double

      Dim YCenter As Double

      Dim thickness As Double

      Dim w As Double

      Dim Rotation As Double



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



   'add new section designer frame section property

      ret = SapModel.PropFrame.SetSDSection("SD1", "A992Fy50")



   'add Plate shape to new property

      ret = SapModel.PropFrame.SDShape.SetPlate("SD1", "SH1", "A992Fy50", 0, -9, 0, -1, 2, 6)



   'get Plate shape property data

      ret = SapModel.PropFrame.SDShape.GetPlate("SD1", "SH1", MatProp, Color, Xcenter, Ycenter, thickness, w, Rotation)



   'close Sap2000

      SapObject.ApplicationExit False

      Set SapModel = Nothing

      Set SapObject = Nothing

End Sub
```



## Release Notes

Initial release in version 12.00.

## See Also

[SetPlate](SetPlate_%7BSDShape%7D.md)

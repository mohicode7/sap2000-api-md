# SetPlate

## Syntax

SapObject.SapModel.PropFrame.SDShape.SetPlate

## VB6 Procedure

Function SetPlate(ByVal Name As String, ByRef ShapeName As String, ByVal MatProp As String, ByVal XCenter As Double, ByVal YCenter As Double, ByVal Rotation As Double, Optional ByVal Color As Long = -1, Optional ByVal thickness As Double = 2.4, Optional ByVal w As Double = 24) As Long

## Parameters

Name

The name of an existing frame section property that is a section designer section.

ShapeName

The name of an existing or new shape in a section designer property. If this is an existing shape,n that shape is modified; otherwise, a new shape is added.

This item may be input as a blank string, in which case the program will assign a shape name to the shape and return that name in the ShapeName variable.

MatProp

The name of the material property for the shape.

XCenter

The X-coordinate of the center of the shape in the section designer coordinate system. [L]

YCenter

The Y-coordinate of the center of the shape in the section designer coordinate system. [L]

Rotation

The counter clockwise rotation of the shape from its default orientation. [deg]

Color

The fill color assigned to the shape. If Color is specified as -1, the program will automatically assign the default fill color.

thickness

The thickness of the plate. [L]

w

The width of the Plate. [L]

## Remarks

This function adds a new Plate shape or modifies an existing shape to be a Plate shape in a section designer property.

The function returns zero if the shape is successfully added or modified; otherwise it returns a nonzero value.

## VBA Example



```
Sub SetFrameSDPropPlate()

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

      ret = SapModel.File.New2DFrame(PortalFrame, 2, 144, 2, 288)



   'add new section designer frame section property

      ret = SapModel.PropFrame.SetSDSection("SD1", "A992Fy50")



   'add Plate shape to new property

      ret = SapModel.PropFrame.SDShape.SetPlate("SD1", "SH1", "A992Fy50", 0, -9, 0, -1, 2, 6)



   'close Sap2000

      SapObject.ApplicationExit False

      Set SapModel = Nothing

      Set SapObject = Nothing

End Sub
```



## Release Notes

Initial release in version 12.00.

## See Also

[GetPlate](GetPlate_%7BSDShape%7D.md)

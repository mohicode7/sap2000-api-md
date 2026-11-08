# SetPipe

## Syntax

SapObject.SapModel.PropFrame.SDShape.SetPipe

## VB6 Procedure

Function SetPipe(ByVal Name As String, ByRef ShapeName As String, ByVal MatProp As String, ByVal PropName As String, ByVal XCenter As Double, ByVal YCenter As Double, Optional ByVal Color As Long = -1, Optional ByVal diameter As Double = 24, Optional ByVal thickness As Double = 2.4) As Long

## Parameters

Name

The name of an existing frame section property that is a section designer section.

ShapeName

The name of an existing or new shape in a section designer property. If this is an existing shape, that shape is modified; otherwise, a new shape is added.

This item may be input as a blank string, in which case the program will assign a shape name to the shape and return that name in the ShapeName variable.

MatProp

The name of the material property for the shape.

PropName

This is a blank string or the name of a defined Pipe property that has been imported from a section property file.

If this item is a blank string, the section dimensions are taken from the values input in this function.

If this item is the name of a defined Pipe property that has been imported from a section property file, the section dimensions are taken from the specified Pipe property.

If this item is not blank and the specified property name is not a Pipe that was imported from a section property file, an error is returned.

XCenter

The X-coordinate of the center of the shape in the section designer coordinate system. [L]

YCenter

The Y-coordinate of the center of the shape in the section designer coordinate system. [L]

Color

The fill color assigned to the shape. If Color is specified as -1, the program will automatically assign the default fill color.

diameter

The outer diameter of the Pipe. [L]

thickness

The wall thickness of the Pipe. [L]

## Remarks

This function adds a new Pipe shape or modifies an existing shape to be a Pipe shape in a section designer property.

The function returns zero if the shape is successfully added or modified; otherwise it returns a nonzero value.

## VBA Example



```
Sub SetFrameSDPropPipe()

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



   'add Pipe shape to new property

      ret = SapModel.PropFrame.SDShape.SetPipe("SD1", "SH1", "A992Fy50", "", 0, -9, -1, 12, 1)



   'close Sap2000

      SapObject.ApplicationExit False

      Set SapModel = Nothing

      Set SapObject = Nothing

End Sub
```



## Release Notes

Initial release in version 12.00.

## See Also

[GetPipe](GetPipe_%7BSDShape%7D.md)

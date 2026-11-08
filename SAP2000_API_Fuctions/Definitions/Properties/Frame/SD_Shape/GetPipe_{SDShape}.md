# GetPipe

## Syntax

SapObject.SapModel.PropFrame.SDShape.GetPipe

## VB6 Procedure

Function GetPipe(ByVal Name As String, ByVal ShapeName As String, ByRef MatProp As String, ByRef PropName As String, ByRef Color As Long, ByRef XCenter As Double, ByRef YCenter As Double, ByRef diameter As Double, ByRef thickness As Double) As Long

## Parameters

Name

The name of an existing frame section property that is a section designer section.

ShapeName

The name of an existing Pipe shape in the specified frame section property.

MatProp

The name of the material property for the shape.

PropName

This is a blank string or the name of a defined Pipe property that has been imported from a section property file. If it is the name of a defined Pipe property, the section dimensions are taken from that property.

Color

The fill color assigned to the shape.

XCenter

The X-coordinate of the center of the shape in the section designer coordinate system. [L]

YCenter

The Y-coordinate of the center of the shape in the section designer coordinate system. [L]

diameter

The outer diameter of the Pipe. [L]

thickness

The wall thickness of the Pipe. [L]

## Remarks

This function retrieves property data for a Pipe shape in a section designer section.

The function returns zero if the property data is successfully retrieved; otherwise it returns a nonzero value.

## VBA Example



```
Sub GetFrameSDPropPipe()

   'dimension variables

      Dim SapObject as cOAPI

      Dim SapModel As cSapModel

      Dim ret As Long

      Dim MatProp As String

      Dim PropName As String

      Dim Color As Long

      Dim XCenter As Double

      Dim YCenter As Double

      Dim diameter As Double

      Dim thickness As Double



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



   'get Pipe shape property data

      ret = SapModel.PropFrame.SDShape.GetPipe("SD1", "SH1", MatProp, PropName, Color, Xcenter, Ycenter, diameter, thickness)



   'close Sap2000

      SapObject.ApplicationExit False

      Set SapModel = Nothing

      Set SapObject = Nothing

End Sub
```



## Release Notes

Initial release in version 12.00.

## See Also

[SetPipe](SetPipe.md)

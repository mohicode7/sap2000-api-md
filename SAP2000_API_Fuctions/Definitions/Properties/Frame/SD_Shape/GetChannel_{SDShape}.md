# GetChannel

## Syntax

SapObject.SapModel.PropFrame.SDShape.GetChannel

## VB6 Procedure

Function GetChannel(ByVal Name As String, ByVal ShapeName As String, ByRef MatProp As String, ByRef PropName As String, ByRef Color As Long, ByRef XCenter As Double, ByRef YCenter As Double, ByRef h As Double, ByRef bf As Double, ByRef tf As Double, ByRef tw As Double, ByRef Rotation As Double) As Long

## Parameters

Name

The name of an existing frame section property that is a section designer section.

ShapeName

The name of an existing Channel shape in the specified frame section property.

MatProp

The name of the material property for the shape.

PropName

This is a blank string or the name of a defined Channel property that has been imported from a section property file. If it is the name of a defined Channel property, the section dimensions are taken from that property.

Color

The fill color assigned to the shape.

XCenter

The X-coordinate of the center of the shape in the section designer coordinate system. [L]

YCenter

The Y-coordinate of the center of the shape in the section designer coordinate system. [L]

h

The section depth. [L]

bf

The flange width. [L]

tf

The flange thickness. [L]

tw

The web thickness. [L]

Rotation

The counter clockwise rotation of the shape from its default orientation. [deg]

## Remarks

This function retrieves property data for a Channel shape in a section designer section.

The function returns zero if the property data is successfully retrieved; otherwise it returns a nonzero value.

## VBA Example



```
Sub GetFrameSDPropChannel()

   'dimension variables

      Dim SapObject as cOAPI

      Dim SapModel As cSapModel

      Dim ret As Long

      Dim MatProp As String

      Dim PropName As String

      Dim Color As Long

      Dim XCenter As Double

      Dim YCenter As Double

      Dim h As Double

      Dim bf As Double

      Dim tf As Double

      Dim tw As Double

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



   'add Channel shape to new property

      ret = SapModel.PropFrame.SDShape.SetChannel("SD1", "SH1", "A992Fy50", "", 0, -9, 0, -1, 18, 6, 1, 0.5)



   'get Channel shape property data

      ret = SapModel.PropFrame.SDShape.GetChannel("SD1", "SH1", MatProp, PropName, Color, Xcenter, Ycenter, h, bf, tf, tw, Rotation)



   'close Sap2000

      SapObject.ApplicationExit False

      Set SapModel = Nothing

      Set SapObject = Nothing

End Sub
```



## Release Notes

Initial release in version 12.00.

## See Also

[SetChannel](SetChannel_%7BSDShape%7D.md)

# GetTee

## Syntax

SapObject.SapModel.PropFrame.SDShape.GetTee

## VB6 Procedure

Function GetTee(ByVal Name As String, ByVal ShapeName As String, ByRef MatProp As String, ByRef PropName As String, ByRef Color As Long, ByRef XCenter As Double, ByRef YCenter As Double, ByRef h As Double, ByRef bf As Double, ByRef tf As Double, ByRef tw As Double, ByRef Rotation As Double) As Long

## Parameters

Name

The name of an existing frame section property that is a section designer section.

ShapeName

The name of an existing Tee shape in the specified frame section property.

MatProp

The name of the material property for the shape.

PropName

This is a blank string or the name of a defined Tee property that has been imported from a section property file. If it is the name of a defined Tee property, the section dimensions are taken from that property.

Color

The fill color assigned to the shape.

XCenter

The X-coordinate of the center of the shape in the section designer coordinate system. [L]

YCenter

The Y-coordinate of the center of the shape in the section designer coordinate system. [L]

h

The section depth. [L]

bf

The section width. [L]

tf

The flange thickness. [L]

tw

The web thickness. [L]

Rotation

The counter clockwise rotation of the shape from its default orientation. [deg]

## Remarks

This function retrieves property data for a Tee shape in a section designer section.

The function returns zero if the property data is successfully retrieved; otherwise it returns a nonzero value.

## VBA Example



```
Sub GetFrameSDPropTee()

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



   'add Tee shape to new property

      ret = SapModel.PropFrame.SDShape.SetTee("SD1", "SH1", "A992Fy50", "", 0, -9, 0, -1, 18, 6, 1, 0.5)



   'get Tee shape property data

      ret = SapModel.PropFrame.SDShape.GetTee("SD1", "SH1", MatProp, PropName, Color, Xcenter, Ycenter, h, bf, tf, tw, Rotation)



   'close Sap2000

      SapObject.ApplicationExit False

      Set SapModel = Nothing

      Set SapObject = Nothing

End Sub
```



## Release Notes

Initial release in version 12.00.

## See Also

[SetTee](SetTee_%7BSDShape%7D.md)

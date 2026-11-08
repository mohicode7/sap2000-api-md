# GetDblAngle

## Syntax

SapObject.SapModel.PropFrame.SDShape.GetDblAngle

## VB6 Procedure

Function GetDblAngle(ByVal Name As String, ByVal ShapeName As String, ByRef MatProp As String, ByRef PropName As String, ByRef Color As Long, ByRef XCenter As Double, ByRef YCenter As Double, ByRef h As Double, ByRef w As Double, ByRef tf As Double, ByRef tw As Double, ByRef dis As Double, ByRef Rotation As Double) As Long

## Parameters

Name

The name of an existing frame section property that is a section designer section.

ShapeName

The name of an existing Double Angle shape in the specified frame section property.

MatProp

The name of the material property for the shape.

PropName

This is a blank string or the name of a defined Double Angle property that has been imported from a section property file. If it is the name of a defined Double Angle property, then the section dimensions are taken from that property.

Color

The fill color assigned to the shape.

XCenter

The X-coordinate of the center of the shape in the section designer coordinate system. [L]

YCenter

The Y-coordinate of the center of the shape in the section designer coordinate system. [L]

h

The section depth. [L]

w

The flange width. [L]

tf

The flange thickness. [L]

tw

The web thickness. [L]

dis

Separation between the two flanges that are parallel. [L]

Rotation

The counter clockwise rotation of the shape from its default orientation. [deg]

## Remarks

This function retrieves property data for a Double Angle shape in a section designer section.

The function returns zero if the property data is successfully retrieved; otherwise it returns a nonzero value.

## VBA Example



```
Sub GetFrameSDPropDblAngle()

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

      Dim w As Double

      Dim tf As Double

      Dim tw As Double

      Dim dis As Double

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



   'add Double Angle shape to new property

      ret = SapModel.PropFrame.SDShape.SetDblAngle("SD1", "SH1", "A992Fy50", "", 0, -9, 0, -1, 18, 6, 1, 0.5, 2)



   'get Double Angle shape property data

      ret = SapModel.PropFrame.SDShape.GetDblAngle("SD1", "SH1", MatProp, PropName, Color, Xcenter, Ycenter, h, w, tf, tw, dis,  Rotation)



   'close Sap2000

      SapObject.ApplicationExit False

      Set SapModel = Nothing

      Set SapObject = Nothing

End Sub
```



## Release Notes

Initial release in version 12.00.

## See Also

[SetDblAngle](SetDblAngle_%7BSDShape%7D.md)

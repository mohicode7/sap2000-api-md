# SetTee

## Syntax

SapObject.SapModel.PropFrame.SDShape.SetTee

## VB6 Procedure

Function SetTee(ByVal Name As String, ByRef ShapeName As String, ByVal MatProp As String, ByVal PropName As String, ByVal XCenter As Double, ByVal YCenter As Double, ByVal Rotation As Double, Optional ByVal Color As Long = -1, Optional ByVal h As Double = 24, Optional ByVal bf As Double = 24, Optional ByVal tf As Double = 2.4, Optional ByVal tw As Double = 2.4) As Long

## Parameters

Name

The name of an existing frame section property that is a section designer section.

ShapeName

The name of an existing or new shape in a section designer property. If this is an existing shape, that shape is modified; otherwise, a new shape is added.

This item may be input as a blank string, in which case the program will assign a shape name to the shape and return that name in the ShapeName variable.

MatProp

The name of the material property for the shape.

PropName

This is a blank string or the name of a defined Tee property that has been imported from a section property file.

If this item is a blank string, the section dimensions are taken from the values input in this function.

If this item is the name of a defined Tee property that has been imported from a section property file, the section dimensions are taken from the specified Tee property.

If this item is not blank, and the specified property name is not a Tee that was imported from a section property file, an error is returned.

XCenter

The X-coordinate of the center of the shape in the section designer coordinate system. [L]

YCenter

The Y-coordinate of the center of the shape in the section designer coordinate system. [L]

Rotation

The counter clockwise rotation of the shape from its default orientation. [deg]

Color

The fill color assigned to the shape. If Color is specified as -1, the program will automatically assign the default fill color.

h

The section depth. [L]

bf

The section width. [L]

tf

The flange thickness. [L]

tw

The web thickness. [L]

## Remarks

This function adds a new Tee shape or modifies an existing shape to be a Tee shape in a section designer property.

The function returns zero if the shape is successfully added or modified; otherwise it returns a nonzero value.

## VBA Example



```
Sub SetFrameSDPropTee()

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



   'add Tee shape to new property

      ret = SapModel.PropFrame.SDShape.SetTee("SD1", "SH1", "A992Fy50", "", 0, -9, 0, -1, 18, 6, 1, 0.5)



   'close Sap2000

      SapObject.ApplicationExit False

      Set SapModel = Nothing

      Set SapObject = Nothing

End Sub
```



## Release Notes

Initial release in version 12.00.

## See Also

[GetTee](GetTee_%7BSDShape%7D.md)

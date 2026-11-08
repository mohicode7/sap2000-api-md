# SetReinfRectangular

## Syntax

SapObject.SapModel.PropFrame.SDShape.SetReinfRectangular

## VB6 Procedure

Function SetReinfRectangular(ByVal Name As String, ByRef
ShapeName As String, ByVal XCenter As Double, ByVal YCenter As Double,
ByVal h As Double, ByVal Rotation As Double, ByVal w As Double, Optional
ByVal MatRebar As String = "") As Long

## Parameters

Name

The name of an existing frame section property that
is a section designer section.

ShapeName

The name of an existing or new shape in a section designer
property. If this is an existing shape, that shape is modified; otherwise,
a new shape is added.

This item may be input as a blank string, in which case
the program will assign a shape name to the shape and return that name
in the ShapeName variable.

XCenter

The X-coordinate of the center of the shape in the section
designer coordinate system. [L]

YCenter

The Y-coordinate of the center of the shape in the section
designer coordinate system. [L]

h

The section depth. [L]

w

The top flange width. [L]

Rotation

The counter clockwise rotation of the shape from its
default orientation. [deg]

MatRebar

The material property for the reinforcing steel.

## Remarks

This function adds a new rectangular reinforcing shape
or modifies an existing shape to be a rectangular reinforcing shape in
a section designer property.

The function returns zero if the shape is successfully
added or modified; otherwise it returns a nonzero value.

## VBA Example



```
Sub SetFrameSDPropReinfRectangular()

   'dimension variables

      Dim SapObject as cOAPI

      Dim SapModel As cSapModel

      Dim ret As Long

      Dim Name As String



   'create Sap2000 object

      Set SapObject = CreateObject("CSI.SAP2000.API.SapObject")



   'start Sap2000 application

      SapObject.ApplicationStart



   'create SapModel object

      Set SapModel = SapObject.SapModel



   'initialize model

      ret = SapModel.InitializeNewModel



   'create model from template

      ret = SapModel.File.New2DFrame(PortalFrame, 
 2, 144, 2, 288)



   'add ASTM A706 rebar material

      ret = SapModel.PropMaterial.AddQuick(Name, 
 MATERIAL_REBAR, , , , , MATERIAL_REBAR_SUBTYPE_ASTM_A706)



   'add new section designer frame section property

      ret = SapModel.PropFrame.SetSDSection("SD1", 
 "4000psi")



   'add rectangular reinforcing shape to new property

      ret = SapModel.PropFrame.SDShape.SetReinfRectangular("SD1", 
 "SH1", 0, 0, 12, 12, 45, Name)



   'close Sap2000

      SapObject.ApplicationExit False

      Set SapModel = Nothing

      Set SapObject = Nothing

End Sub
```



## Release Notes

Initial release in version 12.00.

Moved Rotation to after the width, w, in version 25.0.0

## See Also

[GetReinfRectangular](GetReinfRectangular.md)

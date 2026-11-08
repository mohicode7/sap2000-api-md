# SetSolidSector

## Syntax

SapObject.SapModel.PropFrame.SDShape.SetSolidSector

## VB6 Procedure

Function SetSolidSector(ByVal Name As String, ByRefShapeName As String, ByValMatProp As String, ByValXCenter As Double, ByValYCenter As Double, ByVal Angle As Double, ByVal Rotation As Double, ByValRadius As Double, Optional ByVal Color As Long = -1) As Long

## Parameters

Name

The name of an existing frame section property that is a section designer section.

ShapeName

The name of an existing or new shape in a section designer property. If this is an existing shape, that shape is modified; otherwise, a new shape is added.

This item may be input as a blank string, in which case the program will assign a shape name to the shape and return that name in the ShapeName variable.

MatProp

The name of the material property for the shape.

XCenter

The X-coordinate of the center of the shape in the section designer coordinate system. [L]

YCenter

The Y-coordinate of the center of the shape in the section designer coordinate system. [L]

Angle

The angle between the two radii that define the circular sector. [deg]

Rotation

The counter clockwise rotation of the shape from its default orientation. [deg]

Radius

The radius of the circle defining the Sector. [L]

Color

The fill color assigned to the shape. If Color is specified as -1, the program will automatically assign the default fill color.

## Remarks

This function adds a new solid sector shape or modifies an existing shape to be a solid sector shape in a section designer property.

The function returns zero if the shape is successfully added or modified; otherwise it returns a nonzero value.

## VBA Example



```
Sub SetFrameSDPropSolidSector()

   'dimension variables

      Dim SapObject as cOAPI

      Dim SapModel As cSapModel

      Dim ret As Long

      Dim Name As String



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



   'addASTMA706 rebar material

      ret = SapModel.PropMaterial.AddQuick(Name, MATERIAL_REBAR, , , , , MATERIAL_REBAR_SUBTYPE_ASTM_A706)



   'add new section designer frame section property

      ret = SapModel.PropFrame.SetSDSection("SD1", "4000Psi")



   'add solid sector shape to new property

      ret = SapModel.PropFrame.SDShape.SetSolidSector("SD1", "SH1", "4000Psi", 0, 0, 95, 0, 10, -1)



   'close Sap2000

      SapObject.ApplicationExit False

      Set SapModel = Nothing

      Set SapObject = Nothing

End Sub
```



## Release Notes

Initial release in version 12.00.

## See Also

[GetSolidSector](GetSolidSector.md)

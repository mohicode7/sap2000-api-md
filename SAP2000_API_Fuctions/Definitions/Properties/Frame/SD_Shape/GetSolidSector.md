# GetSolidSector

## Syntax

SapObject.SapModel.PropFrame.SDShape.GetSolidSector

## VB6 Procedure

Function GetSolidSector(ByVal Name As String, ByValShapeName As String, ByRefMatProp As String, ByRef Color As Long, ByRefXCenter As Double, ByRefYCenter As Double, ByRef Angle As Double, ByRefRotation As Double, ByRef Radius As Double) As Long

## Parameters

Name

The name of an existing frame section property that is a section designer section.

ShapeName

The name of an existing solid sector shape in the specified frame section property.

MatProp

The name of the material property for the shape.

Color

The fill color assigned to the shape.

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

## Remarks

This function retrieves property data for a solid sector shape in a section designer section.

The function returns zero if the property data is successfully retrieved; otherwise it returns a nonzero value.

## VBA Example



```
Sub GetFrameSDPropSolidSector()

   'dimension variables

      Dim SapObject as cOAPI

      Dim SapModel As cSapModel

      Dim ret As Long

      Dim Name As String

      Dim MatProp As String

      Dim Color As Long

      Dim XCenter As Double

      Dim YCenter As Double

      Dim Angle As Double

      Dim Rotation As Double

      Dim Radius As Double



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



   'get solid sector shape property data

      ret = SapModel.PropFrame.SDShape.GetSolidSector("SD1", "SH1", MatProp, Color, Xcenter, Ycenter, Angle, Rotation, Radius)



   'close Sap2000

      SapObject.ApplicationExit False

      Set SapModel = Nothing

      Set SapObject = Nothing

End Sub
```



## Release Notes

Initial release in version 12.00.

## See Also

[SetSolidSector](SetSolidSector.md)

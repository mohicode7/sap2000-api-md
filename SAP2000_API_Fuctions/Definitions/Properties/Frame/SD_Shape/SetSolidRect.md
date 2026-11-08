# SetSolidRect

## Syntax

SapObject.SapModel.PropFrame.SDShape.SetSolidRect

## VB6 Procedure

Function SetSolidRect(ByVal Name As String, ByRef ShapeName As String, ByVal MatProp As String, ByVal SSOverwrite As String, ByVal XCenter As Double, ByVal YCenter As Double, ByVal h As Double, ByVal w As Double, ByVal Rotation As Double, Optional ByVal color As Long = -1, Optional ByVal Reinf As Boolean = False, Optional ByVal MatRebar As String = "") As Long

## Parameters

Name

The name of an existing frame section property that is a section designer section.

ShapeName

The name of an existing or new shape in a section designer property. If this is an existing shape, that shape is modified; otherwise, a new shape is added.

This item may be input as a blank string, in which case the program will assign a shape name to the shape and return that name in the ShapeName variable.

MatProp

The name of the material property for the shape.

SSOverwrite

This is a blank string, Default, or the name of a defined stress-strain curve.

If this item is a blank string or Default, the shape stress-strain curve is based on the assigned material property.

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

w

The section width. [L]

Rotation

The counter clockwise rotation of the shape from its default orientation. [deg]

Color

The fill color assigned to the shape. If Color is specified as -1, the program will automatically assign the default fill color.

Reinf

This item is True when there is edge and corner reinforcing steel associated with the shape.

If the MatProp item is not a concrete material, this item is always assumed to be False.

MatRebar

The material property for the edge and corner reinforcing steel associated with the shape. This item applies only when the MatProp item is a concrete material and the Reinf item is True.

## Remarks

This function adds a new solid rectangle shape or modifies an existing shape to be an solid rectangle shape in a section designer property.

The function returns zero if the shape is successfully added or modified; otherwise it returns a nonzero value.

## VBA Example



```
Sub SetFrameSDPropSolidRect()

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

      ret = SapModel.File.New2DFrame(PortalFrame, 2, 144, 2, 288)



   'add ASTM A706 rebar material

      ret = SapModel.PropMaterial.AddQuick(Name, MATERIAL_REBAR, , , , , MATERIAL_REBAR_SUBTYPE_ASTM_A706)



   'add new section designer frame section property

      ret = SapModel.PropFrame.SetSDSection("SD1", "A992Fy50")



   'add solid rectangle shape to new property

      ret = SapModel.PropFrame.SDShape.SetSolidRect("SD1", "SH1", "4000Psi", "Default", 0, 0, 24, 16, 0, -1, True, Name)



   'close Sap2000

      SapObject.ApplicationExit False

      Set SapModel = Nothing

      Set SapObject = Nothing

End Sub
```



## Release Notes

Initial release in version 12.00.

## See Also

[GetSolidRect](GetSolidRect.md)

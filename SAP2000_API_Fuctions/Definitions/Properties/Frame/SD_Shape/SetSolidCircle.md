# SetSolidCircle

## Syntax

SapObject.SapModel.PropFrame.SDShape.SetSolidCircle

## VB6 Procedure

Function SetSolidCircle(ByVal Name As String, ByRefShapeNameAs String, ByValMatProp As String, ByValSSOverwrite As String, ByValXCenter As Double, ByValYCenter As Double, Optional ByVal diameter As Double = 24, Optional ByVal color As Long = -1, Optional ByValReinf As Boolean = False, Optional ByValNumberBars As Long = 8, Option ByVal Rotation As Double = 22.5, Optional ByVal Cover As Double = 2, Optional ByValRebarSize As String = "", Optional ByValMatRebar As String = "") As Long

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

Diameter

The diameter of the circle.[L]

Color

The fill color assigned to the shape. If Color is specified as -1, the program will automatically assign the default fill color.

Reinf

This item is True when there is edge and corner reinforcing steel associated with the shape.

If the MatProp item is not a concrete material, this item is always assumed to be False.

# Bars

This item is visible only if the Reinforcing item is set to True.  It is the number of equally spaced bars for the circular reinforcing.

Rotation

The counter clockwise rotation of the shape from its default orientation. [deg]

Bar Cover

This item is visible only if the Reinforcing item is set to True.  It is the clear cover for the specified rebar. [L]

Barsize

This item is visible only if the Reinforcing item is set to True.  It is the size of the reinforcing bar.

MatRebar

The material property for the edge and corner reinforcing steel associated with the shape. This item applies only when the MatProp item is a concrete material and the Reinf item is True.

## Remarks

This function adds a new solid circle shape or modifies an existing shape to be a solid circle shape in a section designer property.

The function returns zero if the shape is successfully added or modified; otherwise it returns a nonzero value.

## VBA Example



```
Sub SetFrameSDPropSolidCircle()

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



   'add solid circle shape to new property

      ret = SapModel.PropFrame.SDShape.SetSolidCircle("SD1", "SH1", "4000psi", "Default", 0, 0, 12, -1, True, 16, 0, 16, "#4", Name)



   'close Sap2000

      SapObject.ApplicationExit False

      Set SapModel = Nothing

      Set SapObject = Nothing

End Sub
```



## Release Notes

Initial release in version 12.00.

## See Also

[GetSolidCircle](GetSolidCircle.md)

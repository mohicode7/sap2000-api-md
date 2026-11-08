# GetSolidCircle

## Syntax

SapObject.SapModel.PropFrame.SDShape.GetSolidCircle

## VB6 Procedure

Function GetSolidCircle(ByVal Name As String, ByValShapeName As String, ByRefMatProp As String, ByRefSSOverwrite As String, ByRef Color As Long, ByRefXCenter As Double, ByRefYCenter As Double, Optional ByRefDiameter As Double, ByRefReinfAs Boolean, ByRefNumberBarsAs Long = 8, ByRef Rotation As Double, ByRef Cover As Double, ByRefRebarSize As String, ByRefMatRebar As String) As Long

## Parameters

Name

The name of an existing frame section property that is a section designer section.

ShapeName

The name of an existing solid circle shape in the specified frame section property.

MatProp

The name of the material property for the shape.

SSOverwrite

This is a blank string, Default, or the name of a defined stress-strain curve.

If this item is a blank string or Default, the shape stress-strain curve is based on the assigned material property.

Color

The fill color assigned to the shape.

XCenter

The X-coordinate of the center of the shape in the section designer coordinate system. [L]

YCenter

The Y-coordinate of the center of the shape in the section designer coordinate system. [L]

Diameter

The diameter of the circle.[L]

Reinf

This item is True when there is edge and corner reinforcing steel associated with the shape. The MatProp item must refer to a concrete material for this item to be True.

# Bars

This item is visible only if the Reinforcing item is set to True.  It is the number of equally spaced bars for the circular reinforcing.

Rotation

The counter clockwise rotation of the shape from its default orientation. [deg]

Barcover

This item is visible only if the Reinforcing item is set to True. It is the clear cover for the specified rebar.

Bar Size

This item is visible only if the Reinforcing item is set to True. It is the size of the reinforcing bar.

MatRebar

The material property for the edge and corner reinforcing steel associated with the shape. This item applies only when the MatProp item is a concrete material and the Reinf item is True.

## Remarks

This function retrieves property data for a solid circle shape in a section designer section.

The function returns zero if the property data is successfully retrieved; otherwise it returns a nonzero value.

## VBA Example



```
Sub GetFrameSDPropSolidCircle()

   'dimension variables

      Dim SapObject as cOAPI

      Dim SapModel As cSapModel

      Dim ret As Long

      Dim Name As String

      Dim MatProp As String

      Dim SSOverwrite As String

      Dim Color As Long

      Dim XCenter As Double

      Dim YCenter As Double

      Dim Diameter As Long

      Dim Reinf As Boolean

      Dim NumberBars As Long

      Dim Rotation As Double

      Dim Cover As Double

      Dim RebarSize As String

      Dim MatRebar As String



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



   'get solid circle shape property data

      ret = SapModel.PropFrame.SDShape.GetSolidCircle("SD1", "SH1", MatProp, SSOverwrite, Color, XCenter, YCenter, Diameter, Reinf, NumberBars, Rotation, Cover, RebarSize, MatRebar)



   'close Sap2000

      SapObject.ApplicationExit False

      Set SapModel = Nothing

      Set SapObject = Nothing

End Sub
```



## Release Notes

Initial release in version 12.00.

## See Also

[SetSolidCircle](SetSolidCircle.md)

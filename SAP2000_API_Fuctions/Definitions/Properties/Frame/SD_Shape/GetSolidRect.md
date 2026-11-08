# GetSolidRect

## Syntax

SapObject.SapModel.PropFrame.SDShape.GetSolidRect

## VB6 Procedure

Function GetSolidRect(ByVal Name As String, ByVal ShapeName As String, ByRef MatProp As String, ByRef SSOverwrite As String, ByRef Color As Long, ByRef XCenter As Double, ByRef YCenter As Double, ByRef h As Double, ByRef w As Double, ByRef Rotation As Double, ByRef Reinf As Boolean, ByRef MatRebar As String) As Long

## Parameters

Name

The name of an existing frame section property that is a section designer section.

ShapeName

The name of an existing solid rectangle shape in the specified frame section property.

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

h

The section depth. [L]

w

The section width. [L]

Rotation

The counter clockwise rotation of the shape from its default orientation. [deg]

Reinf

This item is True when there is edge and corner reinforcing steel associated with the shape. The MatProp item must refer to a concrete material for this item to be True.

MatRebar

The material property for the edge and corner reinforcing steel associated with the shape. This item applies only when the MatProp item is a concrete material and the Reinf item is True.

## Remarks

This function retrieves property data for a solid rectangular shape in a section designer section.

The function returns zero if the property data is successfully retrieved; otherwise it returns a nonzero value.

## VBA Example



```
Sub GetFrameSDPropSolidRect()

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

      Dim h As Double

      Dim w As Double

      Dim Rotation As Double

      Dim Reinf As Boolean

      Dim MatRebar As String



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



   'get solid rectangle shape property data

      ret = SapModel.PropFrame.SDShape.GetSolidRect("SD1", "SH1", MatProp, SSOverwrite, Color, Xcenter, Ycenter, h, w, Rotation, Reinf, MatRebar)



   'close Sap2000

      SapObject.ApplicationExit False

      Set SapModel = Nothing

      Set SapObject = Nothing

End Sub
```



## Release Notes

Initial release in version 12.00.

## See Also

[SetSolidRect](SetSolidRect.md)

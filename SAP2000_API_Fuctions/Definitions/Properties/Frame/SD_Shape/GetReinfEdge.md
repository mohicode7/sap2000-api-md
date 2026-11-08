# GetReinfEdge

## Syntax

SapObject.SapModel.PropFrame.SDShape.GetReinfEdge

## VB6 Procedure

Function GetReinfEdge(ByVal Name As String, ByVal ShapeName As String, ByRef NumberItems As Long, ByRef EdgeNum() As Long, ByRef RebarSize() As String, ByRef Spacing() As Double, ByRef Cover() As Double) As Long

## Parameters

Name

The name of an existing frame section property that is a section designer section.

ShapeName

The name of an existing solid rectangle shape in the specified frame section property.

NumberItems

The number of edges in the shape.

EdgeNum

This is an array that includes the edge number in the shape.

RebarSize

This is an array that includes None or the name of a defined rebar, indicating the rebar assignment to the considered edge.

Spacing

This is an array that includes the rebar maximum center-to-center along the considered edge. [L]

Cover

This is an array that includes the rebar clear cover along the considered edge. [L]

## Remarks

This function retrieves edge reinforcing data for solid rectangle, circle, polygon, and rectangular reinforcing shapes in a section designer property.

The function returns zero if the property data is successfully retrieved; otherwise it returns a nonzero value.

## VBA Example



```
Sub GetFrameSDPropReinfEdge()

   'dimension variables

      Dim SapObject as cOAPI

      Dim SapModel As cSapModel

      Dim ret As Long

      Dim Name As String

      Dim NumberItems As Long

      Dim EdgeNum() As Long

      Dim RebarSize() As String

      Dim Spacing() As Double

      Dim Cover() As Double



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



   'specify edge reinforcing

      ret = SapModel.PropFrame.SDShape.SetReinfEdge("SD1", "SH1", 1, "#7", 8, 1.75, True)

      ret = SapModel.PropFrame.SDShape.SetReinfEdge("SD1", "SH1", 1, "#4", 4, 1.5)



   'get edge reinforcing

      ret = SapModel.PropFrame.SDShape.GetReinfEdge("SD1", "SH1", NumberItems, EdgeNum, RebarSize, Spacing, Cover)



   'close Sap2000

      SapObject.ApplicationExit False

      Set SapModel = Nothing

      Set SapObject = Nothing

End Sub
```



## Release Notes

Initial release in version 12.00.

## See Also

[SetReinfEdge](SetReinfEdge.md)

[SetReinfCorner](SetReinfCorner.md)

[GetReinfCorner](GetReinfCorner.md)

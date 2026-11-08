# GetReinfLine

## Syntax

SapObject.SapModel.PropFrame.SDShape.GetReinfLine

## VB6 Procedure

Function GetReinfLine(ByVal Name As String, ByVal ShapeName As String, ByRef X1 As Double, ByRef Y1 As Double, ByRef X2 As Double, ByRef Y2 As Double, ByRef Spacing As Double, ByRef RebarSize As String, ByRef EndBars As Boolean, ByRef MatRebar As String) As Long

## Parameters

Name

The name of an existing frame section property that is a section designer section.

ShapeName

The name of a line reinforcing shape in the section designer section.

X1

The section designer X coordinate of the first drawn end point of the line reinforcing. [L]

Y1

The section designer Y coordinate of the first drawn end point of the line reinforcing. [L]

X2

The section designer X coordinate of the second drawn end point of the line reinforcing. [L]

Y2

The section designer Y coordinate of the second drawn end point of the line reinforcing. [L]

Spacing

The center-to-center spacing of the bars in the line pattern shape. [L]

Bar Size

The size of the reinforcing bars used in the line reinforcing shape.

EndBars

This item is True when there are bars at the end points of the line reinforcing.

MatRebar

The material property for the reinforcing steel.

## Remarks

This function retrieves property data for a line reinforcing shape in a section designer section.

The function returns zero if the property data is successfully retrieved; otherwise it returns a nonzero value.

## VBA Example



```
Sub GetFrameSDPropReinfLine()

   'dimension variables

      Dim SapObject as cOAPI

      Dim SapModel As cSapModel

      Dim ret As Long

      Dim MatRebar As String

      Dim X1 As Double

      Dim Y1 As Double

      Dim X2 As Double

      Dim Y2 As Double

      Dim Spacing As Double

      Dim RebarSize As String

      Dim EndBars As Boolean

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

      ret = SapModel.PropFrame.SetSDSection("SD1", "4000psi")



   'add line reinforcing shape to new property

      ret = SapModel.PropFrame.SDShape.SetReinfLine("SD1", "SH1", 0, 0, 5, 2, 12, "#9", True, Name)



   'get line reinforcing shape property data

      ret = SapModel.PropFrame.SDShape.GetReinfLine("SD1", "SH1", X1, Y1, X2, Y2, Spacing, RebarSize, EndBars, MatRebar)



   'close Sap2000

      SapObject.ApplicationExit False

      Set SapModel = Nothing

      Set SapObject = Nothing

End Sub
```



## Release Notes

Initial release in version 12.00.

## See Also

[SetReinfLine](SetReinfLine.md)

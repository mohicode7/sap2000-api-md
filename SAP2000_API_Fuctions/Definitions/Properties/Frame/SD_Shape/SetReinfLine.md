# SetReinfLine

## Syntax

SapObject.SapModel.PropFrame.SDShape.SetReinfLine

## VB6 Procedure

Function SetReinfLine(ByVal Name As String, ByRef ShapeName As String, ByVal X1 As Double, ByVal Y1 As Double, ByVal X2 As Double, ByVal Y2 As Double, Optional ByVal Spacing As Double = 6, Optional ByVal RebarSize As String = "", Optional ByVal EndBars As Boolean = False, ByVal MatRebar As String) As Long

## Parameters

Name

The name of an existing frame section property that is a section designer section.

ShapeName

The name of an existing or new shape in a section designer property. If this is an existing shape, that shape is modified; otherwise, a new shape is added.

This item may be input as a blank string, in which case the program will assign a shape name to the shape and return that name in the ShapeName variable.

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

This function adds a new line reinforcing shape or modifies an existing shape to be a line reinforcing shape in a section designer property.

The function returns zero if the shape is successfully added or modified; otherwise it returns a nonzero value.

## VBA Example



```
Sub SetFrameSDPropReinfLine()

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

      ret = SapModel.PropFrame.SetSDSection("SD1", "4000psi")



   'add Line Pattern Reinforcing shape to new property

      ret = SapModel.PropFrame.SDShape.SetReinfLine("SD1", "SH1", 0, 0, 5, 2, 12, "#9", True, Name)



   'close Sap2000

      SapObject.ApplicationExit False

      Set SapModel = Nothing

      Set SapObject = Nothing

End Sub
```



## Release Notes

Initial release in version 12.00.

## See Also

[GetReinfLine](GetReinfLine.md)

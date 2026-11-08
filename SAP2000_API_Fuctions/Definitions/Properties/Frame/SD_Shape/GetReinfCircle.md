# GetReinfCircle

## Syntax

SapObject.SapModel.PropFrame.SDShape.GetReinfCircle

## VB6 Procedure

Function GetReinfCircle(ByVal Name As String, ByVal ShapeName As String, ByRef XCenter As Double, ByRef YCenter As Double, ByRef Diameter As Double, ByRef Numberbars As Long, ByRef Rotation As Double, ByRef RebarSize As String, ByRef MatRebar As String) As Long

## Parameters

Name

The name of an existing frame section property that is a section designer section.

ShapeName

The name of a circular reinforcing shape in the section designer section.

XCenter

The X-coordinate of the center of the shape in the section designer coordinate system. [L]

YCenter

The Y-coordinate of the center of the shape in the section designer coordinate system. [L]

Diameter

The diameter of the circular shape. [L]

NumberBars

The number of equally spaced bars for the circular reinforcing.

Rotation

The counter clockwise rotation of the shape from its default orientation. [deg]

Barsize

It is the size of the reinforcing bar.

MatRebar

The material property for the reinforcing steel.

## Remarks

This function retrieves property data for a circular reinforcing shape in a section designer section.

The function returns zero if the property data is successfully retrieved; otherwise it returns a nonzero value.

## VBA Example



```
Sub GetFrameSDPropReinfCircle()

   'dimension variables

      Dim SapObject as cOAPI

      Dim SapModel As cSapModel

      Dim ret As Long

      Dim MatRebar As String

      Dim XCenter As Double

      Dim YCenter As Double

      Dim Diameter As Double

      Dim Numberbars As Long

      Dim Rotation As Double

      Dim RebarSize As String

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

      ret = SapModel.PropFrame.SetSDSection("SD1", "4000Psi")



   'add circular reinforcing shape to new property

      ret = SapModel.PropFrame.SDShape.SetReinfCircle("SD1",  "SH1", 0, 0, 12, 4, 45, "#9", Name)



   'get circular reinforcing shape property data

      ret = SapModel.PropFrame.SDShape.GetReinfCircle("SD1", "SH1", Xcenter, Ycenter, Diameter, Numberbars, Rotation, RebarSize, MatRebar)



   'close Sap2000

      SapObject.ApplicationExit False

      Set SapModel = Nothing

      Set SapObject = Nothing

End Sub
```



## Release Notes

Initial release in version 12.00.

## See Also

[SetReinfCircle](SetReinfCircle.md)

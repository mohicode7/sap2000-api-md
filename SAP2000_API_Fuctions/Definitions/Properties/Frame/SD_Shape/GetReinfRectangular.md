# GetReinfRectangular

## Syntax

SapObject.SapModel.PropFrame.SDShape.GetReinfRectangular

## VB6 Procedure

Function GetReinfRectangular(ByVal Name As String, ByVal ShapeName As String, ByRef XCenter As Double, ByRef YCenter As Double, ByRef h As Double, ByRef w As Double, ByRef Rotation As Double, ByRef MatRebar As String) As Long

## Parameters

Name

The name of an existing frame section property that is a section designer section.

ShapeName

The name of a rectangular reinforcing shape in the section designer section.

XCenter

The X-coordinate of the center of the shape in the section designer coordinate system. [L]

YCenter

The Y-coordinate of the center of the shape in the section designer coordinate system. [L]

h

The section depth. [L]

w

The top flange width. [L]

Rotation

The counter clockwise rotation of the shape from its default orientation. [deg]

MatRebar

The material property for the reinforcing steel.

## Remarks

This function retrieves property data for a rectangular reinforcing shape in a section designer section.

The function returns zero if the property data is successfully retrieved; otherwise it returns a nonzero value.

## VBA Example



```
Sub GetFrameSDPropReinfRectangular()

   'dimension variables

      Dim SapObject as cOAPI

      Dim SapModel As cSapModel

      Dim ret As Long

      Dim MatRebar As String

      Dim XCenter As Double

      Dim YCenter As Double

      Dim h As Double

      Dim w As Double

      Dim Rotation As Double

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



   'add rectangular reinforcing shape to new property

      ret = SapModel.PropFrame.SDShape.SetReinfRectangular("SD1", "SH1", 0, 0, 12, 12, 45, Name)



   'get rectangular reinforcing shape property data

      ret = SapModel.PropFrame.SDShape.GetReinfRectangular("SD1", "SH1", Xcenter, Ycenter, h, w, Rotation, MatRebar)



   'close Sap2000

      SapObject.ApplicationExit False

      Set SapModel = Nothing

      Set SapObject = Nothing

End Sub
```



## Release Notes

Initial release in version 12.00.

## See Also

[SetReinfRectangular](SetReinfRectangular.md)

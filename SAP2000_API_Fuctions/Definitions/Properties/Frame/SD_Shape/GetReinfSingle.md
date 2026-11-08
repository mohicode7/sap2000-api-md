# GetReinfSingle

## Syntax

SapObject.SapModel.PropFrame.SDShape.GetReinfSingle

## VB6 Procedure

Function GetReinfSingle(ByVal Name As String, ByVal ShapeName As String, ByRef XCenter As Double, ByRef YCenter As Double, ByRef RebarSize As String, ByRef MatRebar As String) As Long

## Parameters

Name

The name of an existing frame section property that is a section designer section.

ShapeName

The name of a single bar reinforcing shape in the section designer section.

XCenter

The X-coordinate of the center of the shape in the section designer coordinate system. [L]

YCenter

The Y-coordinate of the center of the shape in the section designer coordinate system. [L]

Barsize

The size of the reinforcing bar.

MatRebar

The material property for the reinforcing steel.

## Remarks

This function retrieves property data for a single bar reinforcing shape in a section designer section.

The function returns zero if the property data is successfully retrieved; otherwise it returns a nonzero value.

## VBA Example



```
Sub GetFrameSDPropReinfSingle()

   'dimension variables

      Dim SapObject as cOAPI

      Dim SapModel As cSapModel

      Dim ret As Long

      Dim MatRebar As String

      Dim XCenter As Double

      Dim YCenter As Double

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

      ret = SapModel.PropFrame.SetSDSection("SD1", "4000psi")



   'add single bar reinforcing shape to new property

      ret = SapModel.PropFrame.SDShape.SetReinfSingle("SD1", "SH1", 5, 5, "#9", Name)



   'get single bar reinforcing shape property data

      ret = SapModel.PropFrame.SDShape.GetReinfSingle("SD1", "SH1", XCenter, YCenter, RebarSize, MatRebar)



   'close Sap2000

      SapObject.ApplicationExit False

      Set SapModel = Nothing

      Set SapObject = Nothing

End Sub
```



## Release Notes

Initial release in version 12.00.

## See Also

[SetReinfSingle](GetReinfSingle.md)

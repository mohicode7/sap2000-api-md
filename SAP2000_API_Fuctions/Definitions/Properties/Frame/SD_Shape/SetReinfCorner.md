# SetReinfCorner

## Syntax

SapObject.SapModel.PropFrame.SDShape.SetReinfCorner

## VB6 Procedure

Function SetReinfCorner(ByVal Name As String, ByRef ShapeName As String, ByVal PointNum As Long, ByVal RebarSize As String, Optional ByVal All As Boolean = False) As Long

## Parameters

Name

The name of an existing frame section property that is a section designer section.

ShapeName

The name of an existing solid rectangle, circle or polygon shape in the specified section.

PointNum

An corner point number in the shape. This item is ignored if the All item is True.

RebarSize

This is None or the name of a defined rebar, indicating the rebar assignment to the specified corner.

All

If this item is True, the specified rebar data applies to all corners in the shape.

## Remarks

This function specifies corner reinforcing in solid rectangle, circle and polygon shapes in a section designer property.

The function returns zero if the reinforcing is successfully specified; otherwise it returns a nonzero value.

## VBA Example



```
Sub SetFrameSDPropReinfCorner()

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



   'specify corner reinforcing

      ret = SapModel.PropFrame.SDShape.SetReinfCorner("SD1", "SH1", 1, "#9", True)



   'close Sap2000

      SapObject.ApplicationExit False

      Set SapModel = Nothing

      Set SapObject = Nothing

End Sub
```



## Release Notes

Initial release in version 12.00.

## See Also

[GetReinfCorner](GetReinfCorner.md)

[SetReinfEdge](SetReinfEdge.md)

[GetReinfEdge](GetReinfEdge.md)

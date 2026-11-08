# SetRebarColumn

## Syntax

SapObject.SapModel.PropFrame.SetRebarColumn

## VB6 Procedure

Function SetRebarColumn(ByVal Name As String, ByVal MatPropLong As String, ByVal MatPropConfine As String, ByVal Pattern As Long, ByVal ConfineType As Long, ByVal Cover As Double, ByVal NumberCBars As Long, ByVal NumberR3Bars As Long, ByVal NumberR2Bars As Long, ByVal RebarSize As String, ByVal TieSize As String, ByVal TieSpacingLongit As Double, ByVal Number2DirTieBars As Long, ByVal Number3DirTieBars As Long, ByVal ToBeDesigned As Boolean) As Long

## Parameters

Name

The name of an existing frame section property.

MatPropLong

The name of the rebar material property for the longitudinal rebar.

MatPropConfine

The name of the rebar material property for the confinement rebar.

Pattern

This is either 1 or 2, indicating the rebar configuration.

1 = Rectangular

2 = Circular

For circular frame section properties this item must be 2; otherwise an error is returned.

ConfineType

This is either 1 or 2, indicating the confinement bar type.

1 = Ties

2 = Spiral

This item applies only when Pattern = 2. If Pattern = 1, the confinement bar type is assumed to be ties.

Cover

The clear cover for the confinement steel (ties). In the special case of circular reinforcement in a rectangular column, this is the minimum clear cover. [L]

NumberCBars

This item applies to a circular rebar configuration, Pattern = 2. It is the total number of longitudinal reinforcing bars in the column.

NumberR3Bars

This item applies to a rectangular rebar configuration, Pattern = 1. It is the number of longitudinal bars (including the corner bar) on each face of the column that is parallel to the local 3-axis of the column.

NumberR2Bars

This item applies to a rectangular rebar configuration, Pattern = 1. It is the number of longitudinal bars (including the corner bar) on each face of the column that is parallel to the local 2-axis of the column.

RebarSize

The rebar name for the longitudinal rebar in the column.

TieSize

The rebar name for the confinement rebar in the column.

TieSpacingLongit

The longitudinal spacing of the confinement bars (ties). [L]

Number2DirTieBars

This item applies to a rectangular reinforcing configuration, Pattern = 1. It is the number of confinement bars (tie legs) running in the local 2-axis direction of the column.

Number3DirTieBars

This item applies to a rectangular reinforcing configuration, Pattern = 1. It is the number of confinement bars (tie legs) running in the local 3-axis direction of the column.

ToBeDesigned

If this item is True, the column longitudinal rebar is to be designed; otherwise it is to be checked.

## Remarks

This function assigns column rebar data to frame sections.

The function returns zero if the rebar data is successfully assigned; otherwise it returns a nonzero value.

This function applies only to the following section types. Calling this function for any other type of frame section property returns an error.

SECTION\_RECTANGULAR = 8

SECTION\_CIRCLE = 9

The material assigned to the specified frame section property must be concrete or else this function returns an error.

## VBA Example



```
Sub AssignColumnRebar()

   'dimension variables

      Dim SapObject as cOAPI

      Dim SapModel As cSapModel

      Dim ret As Long

      Dim RebarName As String



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



   'set new frame section property

      ret = SapModel.PropFrame.SetRectangle("R1", "4000Psi", 30, 30)



   'add ASTM A706 rebar material

      ret = SapModel.PropMaterial.AddQuick(RebarName, MATERIAL_REBAR, , , , , MATERIAL_REBAR_SUBTYPE_ASTM_A706)



   'set column rebar data

      ret = SapModel.PropFrame.SetRebarColumn("R1", RebarName, RebarName, 2, 2, 2, 10, 0, 0, "#10", "#5", 4, 0, 0, False)



   'close Sap2000

      SapObject.ApplicationExit False

      Set SapModel = Nothing

      Set SapObject = Nothing

End Sub
```



## Release Notes

Initial release in version 11.02.

## See Also

[GetRebarColumn](GetRebarColumn.md)

# SetRebarBeam

## Syntax

SapObject.SapModel.PropFrame.SetRebarBeam

## VB6 Procedure

Function SetRebarBeam(ByVal Name As String, ByVal MatPropLong As String, ByVal MatPropConfine As String, ByVal CoverTop As Double, ByVal CoverBot As Double, ByVal TopLeftArea As Double, ByVal TopRightArea As Double, ByVal BotLeftArea As Double, ByVal BotRightArea As Double) As Long

## Parameters

Name

The name of an existing frame section property.

MatPropLong

The name of the rebar material property for the longitudinal rebar.

MatPropConfine

The name of the rebar material property for the confinement rebar.

CoverTop

The distance from the top of the beam to the centroid of the top longitudinal reinforcement. [L]

CoverBot

The distance from the bottom of the beam to the centroid of the bottom longitudinal reinforcement. [L]

TopLeftArea

The total area of longitudinal reinforcement at the top left end of the beam. [L2]

TopRightArea

The total area of longitudinal reinforcement at the top right end of the beam. [L2]

BotLeftArea

The total area of longitudinal reinforcement at the bottom left end of the beam. [L2]

BotRightArea

The total area of longitudinal reinforcement at the bottom right end of the beam. [L2]

## Remarks

This function assigns beam rebar data to frame sections.

The function returns zero if the rebar data is successfully assigned; otherwise it returns a nonzero value.

This function applies only to the following section types. Calling this function for any other type of frame section property returns an error.

SECTION\_T = 3

SECTION\_ANGLE = 4

SECTION\_RECTANGULAR = 8

SECTION\_CIRCLE = 9

The material assigned to the specified frame section property must be concrete or this function returns an error.

## VBA Example



```
Sub AssignBeamRebar()

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

      ret = SapModel.PropFrame.SetRectangle("R1", "4000Psi", 20, 12)



   'add ASTM A706 rebar material

      ret = SapModel.PropMaterial.AddQuick(RebarName, MATERIAL_REBAR, , , , , MATERIAL_REBAR_SUBTYPE_ASTM_A706)



   'set beam rebar data

      ret = SapModel.PropFrame.SetRebarBeam("R1", RebarName, RebarName, 3.5, 3, 4.1, 4.2, 4.3, 4.4)



   'close Sap2000

      SapObject.ApplicationExit False

      Set SapModel = Nothing

      Set SapObject = Nothing

End Sub
```



## Release Notes

Initial release in version 11.02.

## See Also

[GetRebarBeam](GetRebarBeam.md)

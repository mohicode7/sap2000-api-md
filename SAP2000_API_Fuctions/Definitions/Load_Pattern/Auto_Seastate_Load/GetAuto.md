# GetAuto

## Syntax

SapObject.SapModel.LoadPatterns.AutoSeastate.GetAuto

## VB6 Procedure

Function GetAuto(ByVal Name As String, ByRef LoadMethod As Long, ByRef CSys As String, ByRef AdjustGravityLat As Boolean, ByRef AdjustGravityLatFactor As Double, ByRef AdjustGravityVert As Boolean, ByRef AdjustGravityVertFactor As Double, ByRef CenterRotation() As Double, ByRef Parameters() As Double, ByRef IgnorePhase As Boolean) As Long

## Parameters

Name

The name of an existing seastate-type load pattern.

LoadMethod

This is one of the following three options defining what parameters are being specified.

1 = Rotation / Translations Vertical Input

2 = Rotations / Translations Full Input

3 = Accelerations / Velocities

CSys

The coordinate system used as a reference for specifying the center of rotation location and the inertia load parameters.

AdjustGravityLat

This item only applies when using LoadMethod 1 or 2. It is True if generated lateral loads should include the effects of the rotated structure, otherwise it is False.

AdjustGravityLatFactor

This item only applies when using LoadMethod 1 or 2. This is a scale factor on the lateral effects generated as a result of the rotated structure.

AdjustGravityVert

This item only applies when using LoadMethod 1 or 2. It is True if generated vertical loads should include the effects of the rotated structure, otherwise it is False.

AdjustGravityVertFactor

This item only applies when using LoadMethod 1 or 2. This is a scale factor on the vertical effects generated as a result of the rotated structure.

CenterRotation

This is an array dimensioned to 2 (3 doubles) that defines the coordinates of the center of rotation, with respect to the selected coordinate system.

Parameters

This is an array of the inertia load parameters, based on the specified LoadMethod, as described below.

If LoadMethod = 1, the following 9 parameters should be input:

Parameters(0) - UZ amplitude

Parameters(1) - UZ period

Parameters(2) - UZ phase

Parameters(3) - RX amplitude

Parameters(4)  - RX period

Parameters(5) - RX phase

Parameters(6) - RY amplitude

Parameters(7) - RY period

Parameters(8) - RY phase

If the LoadMethod = 2, the following 18 parameters should be input:

Parameters(0) - UX amplitude

Parameters(1) - UX period

Parameters(2) - UX phase

Parameters(3) - UY amplitude

Parameters(4)  - UY period

Parameters(5) - UY phase

Parameters(6) - UZ amplitude

Parameters(7) - UZ period

Parameters(8) - UZ phase

Parameters(9) - RX amplitude

Parameters(10) - RX period

Parameters(11) - RX phase

Parameters(12) - RY amplitude

Parameters(13)  - RY period

Parameters(14) - RY phase

Parameters(15) - RZ amplitude

Parameters(16) - RZ period

Parameters(17) - RZ phase

If the LoadMethod = 3, the following 9 parameters should be input:

Parameters(0) - Acceleration UX

Parameters(1) - Acceleration UY

Parameters(2) - Acceleration UZ

Parameters(3) - Acceleration RX

Parameters(4)  - Acceleration RY

Parameters(5) - Acceleration RZ

Parameters(6) - Velocity RX

Parameters(7) - Velocity RY

Parameters(8) - Velocity RZ

IgnorePhase

This item only applies when using LoadMethod 1 or 2. It is True if the input phases should be ignored, otherwise it is False.

## Remarks

This function retrieves auto seastate loading parameters.

The function returns zero if the parameters are successfully retrieved; otherwise it returns a nonzero value.

## VBA Example



```
Sub GetAutoSeastate()

   'dimension variables

      Dim SapObject as cOAPI

      Dim SapModel As cSapModel

      Dim ret As Long
      Dim MyCenter(2) As Double
      Dim MyParams(8) As Double
      Dim LoadMethod As Long
      Dim CSys As String
      Dim AdjustGravityLat As Boolean
      Dim LatScaleFactor As Double
      Dim AdjustGravityVert As Boolean
      Dim VertScaleFactor As Double
      Dim Center() As Double
      Dim Params() As Double
      Dim IgnorePhase As Boolean



   'create Sap2000 object

      Set SapObject = CreateObject("CSI.SAP2000.API.SapObject")



   'start Sap2000 application

      SapObject.ApplicationStart



   'create SapModel object

      Set SapModel = SapObject.SapModel



   'initialize model

      ret = SapModel.InitializeNewModel



   'create model from template

      ret = SapModel.File.New3DFrame(BeamSlab, 2, 144, 3, 336, 2, 432)
 
   'define diaphragm constraints
      ret = SapModel.ConstraintDef.SetDiaphragm("Diaph1", Z)
      ret = SapModel.ConstraintDef.SetDiaphragm("Diaph2", Z)



   'add new load pattern

      ret = SapModel.LoadPatterns.Add("SEASTATE", LTYPE_SEASTATE)
 
   'assign seastate parameters
      MyCenter(0) = 0
      MyCenter(1) = 0
      MyCenter(2) = 0
      MyParams(0) = 0
      MyParams(1) = 0
      MyParams(2) = 0
      MyParams(3) = 0.5
      MyParams(4) = 0
      MyParams(5) = 0
      MyParams(6) = 0
      MyParams(7) = 0
      MyParams(8) = 0
      ret = SapModel.LoadPatterns.AutoSeastate.SetAuto("SEASTATE", 3, "Global", True, 1, True, 1, MyCenter, MyParms, False)
 
   'get seastate parameters
      ret = SapModel.LoadPatterns.AutoSeastate.GetAuto("SEASTATE", LoadMethod, CSys, AdjustGravityLat, LatScaleFactor, AdjustGravityVert, VertScaleFactor, Center, Params, IgnorePhase)



   'close Sap2000

      SapObject.ApplicationExit False

      Set SapModel = Nothing

      Set SapObject = Nothing

End Sub
```



## Release Notes

Initial release in version 17.2.0.

## See Also

[SetAuto](SetAuto.md)

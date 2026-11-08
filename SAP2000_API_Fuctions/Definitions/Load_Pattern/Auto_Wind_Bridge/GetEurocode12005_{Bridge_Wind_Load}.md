# GetEurocode12005

## Syntax

SapObject.SapModel.LoadPatterns.AutoWindBridge.GetEurocode12005

## VB6 Procedure

Function GetEurocode12005(ByVal Name As String, ByRef Superstructure As Boolean, ByRef Substructure As Boolean, ByRef Vertical As Boolean, ByRef GroundElevation As Double, ByRef SuperZProgCalc As Boolean, ByRef SuperstructureZ As Double, ByRef SubZProgCalc As Boolean, ByRef SubstructureZ As Double, ByRef WindSpeed As Double, ByRef Terrain As Long, ByRef Orography As Double, ByRef k1 As Double, ByRef Rho As Double, ByRef Cfx As Double, ByRef CfzUp As Double) As Long

## Parameters

Name

The name of an existing Wind-type load pattern.

Superstructure

This item is True if wind on the superstructure should be considered, otherwise it is false.

Substructure

This item is True if wind on the substructure should be considered, otherwise it is false.

Vertical

This item is True if vertical wind should be considered, otherwise it is false

GroundElevation

The ground elevation used for determining heights used for wind pressure values. [L]

SuperZProgCalc

This item is True if the superstructure height, Z, should be program calculated, otherwise it is false.

SuperstructureZ

The superstructure height Z. This item applies only when SuperProgCalc = False. [L]

SubZProgCalc

This item is True if the substructure height, Z, should be program calculated, otherwise it is false.

SuperstructureZ

The substructure height Z. This item applies only when SubProgCalc = False. [L]

WindSpeed

The basic wind speed, vb, in meters per second.

Terrain

The terrain category.

0 = 0

1 = I

2 = II

3 = III

4 = IV

Orography

The orography factor, Co.

k1

The turbulence factor, k1.

Rho

The air density in kg/m3, Rho.

Cfx

The force coefficient, Cf, in the horizontal direction.

CfzUp

The force coefficient, Cf, in the vertical up direction.

## Remarks

This function retrieves bridge auto wind loading parameters for Eurocode 1-2005.

The function returns zero if the parameters are successfully retrieved; otherwise it returns a nonzero value.

## VBA Example



```
Sub GetBridgeWindEurocode12005()

   'dimension variables

      Dim SapObject as cOAPI

      Dim SapModel As cSapModel

      Dim ret As Long

      
      Dim Superstructure As Boolean

      Dim Substructure As Boolean

      Dim Vertical As Boolean

      Dim GroundElevation As Double

      Dim SuperZProgCalc As Boolean
      Dim SuperstructureZ As Double
      Dim SubZProgCalc As Boolean
      Dim SubstructureZ As Double

      Dim WindSpeed As Double
      Dim Terrain As Long

      Dim Orography As Double

      Dim k1 As Double
      Dim Rho As Double
      Dim Cfx As Double
      Dim CfzUp As Double



   'create Sap2000 object

      Set SapObject = CreateObject("CSI.SAP2000.API.SapObject")



   'start Sap2000 application

      SapObject.ApplicationStart



   'create SapModel object

      Set SapModel = SapObject.SapModel



   'initialize model

      ret = SapModel.InitializeNewModel



   'open existing model containing a bridge object

      ret = SapModel.File.OpenFile(“C:\Temp\BridgeModel.bdb” )



   'add new load pattern

      ret = SapModel.LoadPatterns.Add("WIND", LTYPE_WIND)



   'assign Eurocode 1-2005 parameters

      ret = SapModel.LoadPatterns.AutoWindBridge.SetEurocode12005("WIND", True, True, True, -100, True, 10, True, 10, 35, 2, 1, 1.25, 1.3, 0.9)



   'get Eurocode 1-2005 parameters

      ret = SapModel.LoadPatterns.AutoWindBridge.GetEurocode12005("WIND", Superstructure, Substructure, Vertical, GroundElevation, SuperZProgCalc, SuperstructureZ, SubZProgCalc, SubstructureZ, WindSpeed, Terrain, Orography, k1, Rho, Cfx, CfzUp)



   'close Sap2000

      SapObject.ApplicationExit False

      Set SapModel = Nothing

      Set SapObject = Nothing

End Sub
 
```



## Release Notes

Initial release in v21.0.0.

## See Also

[SetEurocode12005](SetEurocode12005_%7BBridge_Wind_Load%7D.md)

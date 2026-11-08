# SetEurocode12005

## Syntax

SapObject.SapModel.LoadPatterns.AutoWindBridge.SetEurocode12005

## VB6 Procedure

Function SetEurocode12005(ByVal Name As String, ByVal Superstructure As Boolean, ByVal Substructure As Boolean, ByVal Vertical As Boolean, ByVal GroundElevation As Double, ByVal SuperZProgCalc As Boolean, ByVal SuperstructureZ As Double, ByVal SubZProgCalc As Boolean, ByVal SubstructureZ As Double, ByVal WindSpeed As Double, ByVal Terrain As Long, ByVal Orography As Double, ByVal k1 As Double, ByVal Rho As Double, ByVal Cfx As Double, ByVal CfzUp As Double) As Long

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

This function sets bridge auto wind loading parameters for Eurocode 1-2005.

The function returns zero if the parameters are successfully assigned; otherwise it returns a nonzero value.

## VBA Example



```
Sub SetBridgeWindEurocode12005()

   'dimension variables

      Dim SapObject as cOAPI

      Dim SapModel As cSapModel

      Dim ret As Long

      

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



   'close Sap2000

      SapObject.ApplicationExit False

      Set SapModel = Nothing

      Set SapObject = Nothing

End Sub
 
```



## Release Notes

Initial release in v21.0.0.

## See Also

[GetEurocode12005](GetEurocode12005_%7BBridge_Wind_Load%7D.md)

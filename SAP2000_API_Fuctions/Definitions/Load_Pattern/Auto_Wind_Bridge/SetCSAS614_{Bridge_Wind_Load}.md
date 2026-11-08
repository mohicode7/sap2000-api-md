# SetCSAS614 {Bridge Wind Load}

## Syntax

SapObject.SapModel.LoadPatterns.AutoWindBridge.SetCSAS614

## VB6 Procedure

Function SetCSAS614(ByVal Name As String, ByVal Superstructure As Boolean, ByVal Substructure As Boolean, ByVal Vertical As Boolean, ByVal GroundElevation As Double, ByVal SuperZProgCalc As Boolean, ByVal SuperstructureZ As Double, ByVal q As Double, ByVal Cg As Double, ByVal CeProgCalc As Boolean, ByVal Ce As Double, ByVal ChStructure As Double, ByVal ChLiveLoad As Double, ByVal Cv As Double) As Long

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

q

The reference wind pressure, q. [F/L2]

Cg

The gust effect factor.

CeProgCalc

This item is True if the wind exposure factor, Ce, should be program calculated, otherwise it is false.

Ce

The wind exposure factor. This item applies only when CeProgCalc = False.

ChStructure

The Ch factor for wind acting on the structure.

ChLiveLoad

The Ch factor for wind acting on the live load.

Cv

The Cv factor for vertical wind.

## Remarks

This function sets bridge auto wind loading parameters for CSA S6-14.

The function returns zero if the parameters are successfully assigned; otherwise it returns a nonzero value.

## VBA Example



```
Sub SetBridgeWindCSAS614()

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



   'assign CSA S6-14 parameters

      ret = SapModel.LoadPatterns.AutoWindBridge.SetCSAS614("WIND", True, True, True, -100, True, 10, 1, 2, True, 1, 2, 1.2, 1)



   'close Sap2000

      SapObject.ApplicationExit False

      Set SapModel = Nothing

      Set SapObject = Nothing

End Sub
 
```



## Release Notes

Initial release in v21.0.0.

## See Also

[GetCSAS614](GetCSAS614_%7BBridge_Wind_Load%7D.md)

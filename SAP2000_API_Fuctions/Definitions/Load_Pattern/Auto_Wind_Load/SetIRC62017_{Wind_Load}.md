# SetIRC62017

## Syntax

SapObject.SapModel.LoadPatterns.AutoWindBridge.SetIRC62017

## VB6 Procedure

Function
SetIRC62017(ByVal Name As String, ByVal Superstructure As Boolean, ByVal
Substructure As Boolean, ByVal Vertical As Boolean, ByVal GroundElevation
As Double, ByVal SuperZProgCalc As Boolean, ByVal SuperstructureZ As Double,
ByVal SubZProgCalc As Boolean, ByVal SubstructureZ As Double, ByVal vb
As Double, ByVal Terrain As Long, ByVal G As Double, ByVal CDSuper As
Double, ByVal CDSub As Double, ByVal CL As Double) As Long

## Parameters

Name

The name of an existing Wind-type load pattern.

Superstructure

This item is True if wind on the superstructure should
be considered, otherwise it is false.

Substructure

This item is True if wind on the substructure should
be considered, otherwise it is false.

Vertical

This item is True if vertical wind should be considered,
otherwise it is false

GroundElevation

The ground elevation used for determining heights used
for wind pressure values. [L]

SuperZProgCalc

This item is True if the superstructure height, Z, should
be program calculated, otherwise it is false.

SuperstructureZ

The superstructure height Z. This item applies only
when SuperProgCalc = False. [L]

SubZProgCalc

This item is True if the substructure height, Z, should
be program calculated, otherwise it is false.

SubstructureZ

The substructure height Z. This item applies only when
SubProgCalc = False. [L]

vb

The basic wind speed. [m/s]

Terrain

The terrain type.

0
= Plain

1
= Obstructions

G

The gust factor, G.

CDSuper

The drag coefficient for the superstructure, CD.

CDSub

The drag coefficient for the substructure, CD.

CL

The lift coefficient, CL, for vertical wind.

## Remarks

This function sets bridge auto wind loading parameters
for IRC:6-2017.

The function returns zero if the parameters are successfully
assigned; otherwise it returns a nonzero value.

## VBA Example



```
Sub GetWindIRC62017()

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

      ret = SapModel.File.OpenFile(“C:\Temp\BridgeModel.bdb” 
 )



   'add new load pattern

      ret = SapModel.LoadPatterns.Add("WIND", 
 LTYPE_WIND)



   'assign CSA S6-14 parameters

      ret = SapModel.LoadPatterns.AutoWindBridge. 
 SetIRC62017("WIND", True, True, True, -100, True, 10, True, 
 10, 33, 0, 2, 1.5, 1.5, 0.75)



   'close Sap2000

      SapObject.ApplicationExit False

      Set SapModel = Nothing

      Set SapObject = Nothing



 End Sub
```



## Release Notes

Initial release in v23.1.0.

## See Also

[GetIRC62017](GetIRC62017_%7BWind_Load%7D.md)

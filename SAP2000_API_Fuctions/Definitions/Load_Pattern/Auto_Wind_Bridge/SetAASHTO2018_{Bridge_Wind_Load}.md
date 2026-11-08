# SetAASHTO2018

## Syntax

SapObject.SapModel.LoadPatterns.AutoWindBridge.SetAASHTO2018

## VB6 Procedure

Function SetAASHTO2018(ByVal Name As String, ByVal LimitState As Long, ByVal Superstructure As Boolean, ByVal Substructure As Boolean, ByVal Vertical As Boolean, ByVal GroundElevation As Double, ByVal SuperZProgCalc As Boolean, ByVal SuperstructureZ As Double, ByVal SubZProgCalc As Boolean, ByVal SubstructureZ As Double, ByVal WindSpeed As Double, ByVal ExposureCategory As Long, ByVal G As Double, ByVal CdSuper As Double, ByVal CdSub As Double) As Long

## Parameters

Name

The name of an existing Wind-type load pattern.

LimitState

The limit state associated with the wind load.

1 = Strength III

2 = Strength V

3 = Service I

4 = Service IV

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

The wind speed in mph.

Exposure Category

The wind exposure category.

1 = B

2 = C

3 = D

G

The gust effect factor, G.

CdSuper

The drag coefficient for the superstructure.

CdSub

The drag coefficient for the substructure.

## Remarks

This function sets bridge auto wind loading parameters for AASHTO 2018.

The function returns zero if the parameters are successfully assigned; otherwise it returns a nonzero value.

## VBA Example



```
Sub SetBridgeWindAASHTO2018()

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



   'assign AASHTO2018 parameters

      ret = SapModel.LoadPatterns.AutoWindBridge.SetAASHTO2018("WIND", 1, True, True, True, -100, True, 10, True, 10, 100, 2, 1, 1.3, 1.6)



   'close Sap2000

      SapObject.ApplicationExit False

      Set SapModel = Nothing

      Set SapObject = Nothing

End Sub
```



## Release Notes

Initial release in v21.0.0.

## See Also

[GetAASHTO2018](GetAASHTO2018_%7BBridge_Wind_Load%7D.md)

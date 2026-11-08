# GetAASHTO2020

## Syntax

SapObject.SapModel.LoadPatterns.AutoWindBridge.GetAASHTO2020

## VB6 Procedure

Function GetAASHTO2020(ByVal Name As String, ByRef LimitState As Long, ByRef Superstructure As Boolean, ByRef Substructure As Boolean, ByRef Vertical As Boolean, ByRef GroundElevation As Double, ByRef SuperZProgCalc As Boolean, ByRef SuperstructureZ As Double, ByRef SubZProgCalc As Boolean, ByRef SubstructureZ As Double, ByRef WindSpeed As Double, ByRef ExposureCategory As Long, ByRef G As Double, ByRef CdSuper As Double, ByRef CdSub As Double) As Long

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

This function retrieves bridge auto wind loading parameters for AASHTO 2020.

The function returns zero if the parameters are successfully retrieved; otherwise it returns a nonzero value.

## VBA Example



```
Sub GetBridgeWindAASHTO2020()

       'dimension variables

          Dim SapObject as cOAPI

          Dim SapModel As cSapModel

          Dim ret As Long

          
      Dim LimitState As Long

          Dim Superstructure As Boolean

          Dim Substructure As Boolean

          Dim Vertical As Boolean

          Dim GroundElevation As Double

          Dim SuperZProgCalc As Boolean
      Dim SuperstructureZ As Double

          Dim SubZProgCalc As Boolean
      Dim SubstructureZ As Double

          Dim WindSpeed As Double

          Dim ExposureCategory As Long
      Dim G As Double
      Dim CdSuper As Double
      Dim CdSub As Double



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



       'assign AASHTO2020 parameters

          ret = SapModel.LoadPatterns.AutoWindBridge.SetAASHTO2020("WIND", 1, True, True, True, -100, True, 10, True, 10, 100, 2, 1, 1.3, 1.6)



       'get AASHTO 2020 parameters

          ret = SapModel.LoadPatterns.AutoWindBridge.GetAASHTO2020("WIND", LimitState, Superstructure, Substructure, Vertical, GroundElevation, SuperZProgCalc, SuperstructureZ, SubZProgCalc, SubstructureZ, WindSpeed, ExposureCategory, G, CdSuper, CdSub)



       'close Sap2000

          SapObject.ApplicationExit False

          Set SapModel = Nothing

          Set SapObject = Nothing

    End Sub
  
 
```



## Release Notes

Initial release in v25.3.0.

## See Also

[SetAASHTO2020](SetAASHTO2020_%7BBridge_Wind_Load%7D.md)

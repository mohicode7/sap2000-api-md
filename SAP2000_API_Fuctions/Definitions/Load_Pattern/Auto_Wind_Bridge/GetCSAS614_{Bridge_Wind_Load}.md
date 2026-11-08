# GetCSAS614

## Syntax

SapObject.SapModel.LoadPatterns.AutoWindBridge.GetCSAS614

## VB6 Procedure

Function GetCSAS614(ByVal Name As String, ByRef Superstructure As Boolean, ByRef Substructure As Boolean, ByRef Vertical As Boolean, ByRef GroundElevation As Double, ByRef SuperZProgCalc As Boolean, ByRef SuperstructureZ As Double, ByRef q As Double, ByRef Cg As Double, ByRef CeProgCalc As Boolean, ByRef Ce As Double, ByRef ChStructure As Double, ByRef ChLiveLoad As Double, ByRef Cv As Double) As Long

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

This function retrieves bridge auto wind loading parameters for CSA S6-14.

The function returns zero if the parameters are successfully retrieved; otherwise it returns a nonzero value.

## VBA Example



```
Sub GetBridgeWindCSAS614()

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

      Dim q As Double
      Dim Cg As Double

      Dim CeProgCalc As Boolean

      Dim Ce As Double
      Dim ChStructure As Double
      Dim ChLiveLoad As Double
      Dim Cv As Double



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



   'get CSA S6-14 parameters

      ret = SapModel.LoadPatterns.AutoWindBridge.GetCSAS614("WIND", Superstructure, Substructure, Vertical, GroundElevation, SuperZProgCalc, SuperstructureZ, q, Cg, CeProgCalc, Ce, ChStructure, ChLiveLoad, Cv)



   'close Sap2000

      SapObject.ApplicationExit False

      Set SapModel = Nothing

      Set SapObject = Nothing

End Sub
 
```



## Release Notes

Initial release in v21.0.0.

## See Also

[SetCSAS614](SetCSAS614_%7BBridge_Wind_Load%7D.md)

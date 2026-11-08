# SetSelfWtMultiplier

## Syntax

SapObject.SapModel.LoadPatterns.SetSelfWtMultiplier

## VB6 Procedure

Function SetSelfWtMultiplier(ByVal Name As String, ByVal SelfWTMultiplier As Double) As Long

## Parameters

Name

The name of an existing load pattern.

SelfWTMultiplier

The self weight multiplier for the specified load pattern.

## Remarks

This function assigns a self weight multiplier to a load case.

The function returns zero if the multiplier is successfully assigned; otherwise it returns a nonzero value.

## VBA Example



```
Sub AssignSelfWeightMultiplier()

   'dimension variables

      Dim SapObject as cOAPI

      Dim SapModel As cSapModel

      Dim ret As Long



   'create Sap2000 object

      Set SapObject = New Sap2000v16.SapObject



   'start Sap2000 application

      SapObject.ApplicationStart



   'create SapModel object

      Set SapModel = SapObject.SapModel



   'initialize model

      ret = SapModel.InitializeNewModel



   'create model from template

      ret = SapModel.File.New2DFrame(PortalFrame, 2, 144, 2, 288)



   'assign self weight multiplier

      ret = SapModel.LoadPatterns.SetSelfWtMultiplier("DEAD", 2)



   'close Sap2000

      SapObject.ApplicationExit False

      Set SapModel = Nothing

      Set SapObject = Nothing

End Sub
```



## Release Notes

Initial release in version 11.01.

Changed nomenclature from Load Cases, Analysis Cases and Response Combinations to Load Patterns, Load Cases and Load Combinations, respectively, in version 12.00.

## See Also

[GetSelfWtMultiplier](GetSelfWtMultiplier.md)

# GetSelfWtMultiplier

## Syntax

SapObject.SapModel.LoadPatterns.GetSelfWtMultiplier

## VB6 Procedure

Function GetSelfWtMultiplier(ByVal Name As String, ByRef SelfWTMultiplier As Double) As Long

## Parameters

Name

The name of an existing load pattern.

SelfWTMultiplier

The self weight multiplier for the specified load pattern.

## Remarks

This function retrieves the self weight multiplier for a specified load pattern.

The function returns zero if the multiplier is successfully retrieved; otherwise it returns a nonzero value.

## VBA Example



```
Sub GetSelfWeightMultiplier()

   'dimension variables

      Dim SapObject as cOAPI

      Dim SapModel As cSapModel

      Dim ret As Long

      Dim SelfWTMultiplier As Double



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



   'get self weight multiplier

      ret = SapModel.LoadPatterns.GetSelfWtMultiplier("DEAD", SelfWTMultiplier)



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

[SetSelfWtMultiplier](SetSelfWtMultiplier.md)

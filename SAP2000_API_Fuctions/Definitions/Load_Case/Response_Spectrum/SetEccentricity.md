# SetEccentricity

## Syntax

SapObject.SapModel.LoadCases.ResponseSpectrum.SetEccentricity

## VB6 Procedure

Function SetEccentricity(ByVal Name As String, ByVal Eccen As Double) As Long

## Parameters

Name

The name of an existing response spectrum load case.

Eccen

The eccentricity ratio that applies to all diaphragms.

## Remarks

This function sets the eccentricity ratio that applies to all diaphragms for the specified load case.

The function returns zero if the ratio is successfully set; otherwise it returns a nonzero value.

## VBA Example



```
Sub SetCaseResponseSpectrumEccentricity()

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



   'create model from template

      ret = SapModel.File.New2DFrame(PortalFrame, 2, 144, 2, 288)



   'add response spectrum load case

      ret = SapModel.LoadCases.ResponseSpectrum.SetCase("LCASE1")



   'set eccentricity ratio

      ret = SapModel.LoadCases.ResponseSpectrum.SetEccentricity("LCASE1", 0.05)



   'close Sap2000

      SapObject.ApplicationExit False

      Set SapModel = Nothing

      Set SapObject = Nothing

End Sub
```



## Release Notes

Initial release in version 11.02.

Changed nomenclature from Load Cases, Analysis Cases and Response Combinations to Load Patterns, Load Cases and Load Combinations, respectively, in version 12.00.

## See Also

[GetEccentricity](GetEccentricity.md)

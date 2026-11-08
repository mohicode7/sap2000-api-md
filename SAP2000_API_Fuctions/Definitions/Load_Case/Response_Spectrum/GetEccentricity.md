# GetEccentricity

## Syntax

SapObject.SapModel.LoadCases.ResponseSpectrum.GetEccentricity

## VB6 Procedure

Function GetEccentricity(ByVal Name As String, ByRef Eccen As Double) As Long

## Parameters

Name

The name of an existing response spectrum load case.

Eccen

The eccentricity ratio that applies to all diaphragms.

## Remarks

This function retrieves the eccentricity ratio that applies to all diaphragms for the specified load case.

The function returns zero if the ratio is successfully retrieved; otherwise it returns a nonzero value.

## VBA Example



```
Sub GetCaseResponseSpectrumEccentricity()

   'dimension variables

      Dim SapObject as cOAPI

      Dim SapModel As cSapModel

      Dim ret As Long

      Dim Eccen As Double



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



   'get eccentricity ratio

      ret = SapModel.LoadCases.ResponseSpectrum.GetEccentricity("LCASE1", Eccen)



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

[SetEccentricity](SetEccentricity.md)

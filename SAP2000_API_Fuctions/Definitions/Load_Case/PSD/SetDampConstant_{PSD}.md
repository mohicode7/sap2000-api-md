# SetDampConstant

## Syntax

SapObject.SapModel.LoadCases.PSD.SetDampConstant

## VB6 Procedure

Function SetDampConstant(ByVal Name As String, ByVal HysConMassCoeff As Double, ByVal HysConStiffCoeff As Double) As Long

## Parameters

Name

The name of an existing power spectral density analysis case.

HysConMassCoeff

The mass proportional damping coefficient.

HysConStiffCoeff

The stiffness proportional damping coefficient.

## Remarks

This function sets constant hysteretic damping for all frequencies for the specified load case.

The function returns zero if the damping is successfully set; otherwise it returns a nonzero value.

## VBA Example



```
Sub SetCasePSDDampConstant()

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



   'add power spectral density load case

      ret = SapModel.LoadCases.PSD.SetCase("LCASE1")



   'set constant damping

      ret = SapModel.LoadCases.PSD.SetDampConstant("LCASE1", 0.8, 0.04)



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

[GetDampConstant](GetDampConstant_%7BPSD%7D.md)

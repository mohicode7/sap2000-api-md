# GetDampConstant

## Syntax

SapObject.SapModel.LoadCases.PSD.GetDampConstant

## VB6 Procedure

Function GetDampConstant(ByVal Name As String, ByRef HysConMassCoeff As Double, ByRef HysConStiffCoeff As Double) As Long

## Parameters

Name

The name of an existing power spectral density analysis case that has constant damping.

HysConMassCoeff

The mass proportional damping coefficient.

HysConStiffCoeff

The stiffness proportional damping coefficient.

## Remarks

This function retrieves the constant hysteretic damping for all frequencies assigned to the specified load case.

The function returns zero if the damping is successfully retrieved; otherwise it returns a nonzero value.

## VBA Example



```
Sub GetCasePSDDampConstant()

   'dimension variables

      Dim SapObject as cOAPI

      Dim SapModel As cSapModel

      Dim ret As Long

      Dim DampType As Long

      Dim HysConMassCoeff As Double

      Dim HysConStiffCoeff As Double



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



   'get constant damping

      ret = SapModel.LoadCases.PSD.GetDampType("LCASE1", DampType)

      If DampType = 1 Then

         ret = SapModel.LoadCases.PSD.GetDampConstant("LCASE1", HysConMassCoeff, HysConStiffCoeff)

      End If



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

[SetDampConstant](SetDampConstant_%7BPSD%7D.md)

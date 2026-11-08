# GetTimeIntegration

## Syntax

SapObject.SapModel.LoadCases.DirHistNonlinear.GetTimeIntegration

## VB6 Procedure

Function GetTimeIntegration(ByVal Name As String, ByRef IntegrationType As Long, ByRef Alpha As Double, ByRef Beta As Double, ByRef Gamma As Double, ByRef Theta As Double, ByRef m As Double) As Long

## Parameters

Name

The name of an existing nonlinear direct integration time history load case.

IntegrationType

This is 1, 2, 3, 4 or 5; indicating the time integration type.

1 = Newmark

2 = Wilson

3 = Collocation

4 = Hilber-Hughes-Taylor

5 = Chung and Hulbert

Alpha

The alpha factor (-1/3 <= Alpha <= 0).

This item applies only when IntegrationType = 4 or 5.

Beta

The beta factor (Beta >= 0).

This item applies only when IntegrationType = 1, 3 or 5. It is returned for informational purposes when IntegrationType = 4.

Gamma

The gamma factor (Gamma >= 0.5).

This item applies only when IntegrationType = 1, 3 or 5. It is returned for informational purposes when IntegrationType = 4.

Theta

The theta factor (Theta > 0).

This item applies only when IntegrationType = 2 or 3.

m

The alpha-m factor.

This item applies only when IntegrationType = 5.

## Remarks

This function retrieves the time integration data assigned to the specified load case.

The function returns zero if the option is successfully retrieved; otherwise it returns a nonzero value.

## VBA Example



```
Sub GetCaseDirHistNonlinearTimeIntegration()

   'dimension variables

      Dim SapObject as cOAPI

      Dim SapModel As cSapModel

      Dim ret As Long

      Dim IntegrationType As Long

      Dim Alpha As Double

      Dim Beta As Double

      Dim Gamma As Double

      Dim Theta As Double

      Dim m As Double



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



   'add nonlinear direct history load case

      ret = SapModel.LoadCases.DirHistNonlinear.SetCase("LCASE1")



   'set time integration parameters

      ret = SapModel.LoadCases.DirHistNonlinear.SetTimeIntegration("LCASE1", 3, 0, 0.17, 0.52, 0.9)



   'get time integration parameters

      ret = SapModel.LoadCases.DirHistNonlinear.GetTimeIntegration("LCASE1", IntegrationType, Alpha, Beta, Gamma, Theta, m)



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

[SetTimeIntegration](SetTimeIntegration_%7BDirect_TH_Nonlinear%7D.md)

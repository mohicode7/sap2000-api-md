# GetDampConstant

## Syntax

SapObject.SapModel.LoadCases.ModHistNonlinear.GetDampConstant

## VB6 Procedure

Function GetDampConstant(ByVal Name As String, ByRef Damp As Double) As Long

## Parameters

Name

The name of an existing nonlinear modal history analysis case that has constant damping.

Damp

The constant damping for all modes (0 <= Damp < 1).

## Remarks

This function retrieves the constant modal damping for all modes assigned to the specified load case.

The function returns zero if the option is successfully retrieved; otherwise it returns a nonzero value.

## VBA Example



```
Sub GetCaseModHistNonlinearDampConstant()

   'dimension variables

      Dim SapObject as cOAPI

      Dim SapModel As cSapModel

      Dim ret As Long

      Dim DampType As Long

      Dim Damp As Double



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



   'add nonlinear modal history load case

      ret = SapModel.LoadCases.ModHistNonlinear.SetCase("LCASE1")



   'set constant damping

      ret = SapModel.LoadCases.ModHistNonlinear.SetDampConstant("LCASE1", 0.04)



   'get constant damping

      ret = SapModel.LoadCases.ModHistNonlinear.GetDampType("LCASE1", DampType)

      If DampType = 4 Then

         ret = SapModel.LoadCases.ModHistNonlinear.GetDampConstant("LCASE1", Damp)

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

[SetDampConstant](SetDampConstant.md)

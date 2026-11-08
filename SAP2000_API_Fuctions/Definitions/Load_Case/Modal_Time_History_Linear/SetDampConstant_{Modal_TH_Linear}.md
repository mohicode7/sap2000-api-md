# SetDampConstant

## Syntax

SapObject.SapModel.LoadCases.ModHistLinear.SetDampConstant

## VB6 Procedure

Function SetDampConstant(ByVal Name As String, ByVal Damp As Double) As Long

## Parameters

Name

The name of an existing linear modal history analysis case.

Damp

The constant damping for all modes (0 <= Damp < 1).

## Remarks

This function sets constant modal damping for the specified load case.

The function returns zero if the damping is successfully set; otherwise it returns a nonzero value.

## VBA Example



```
Sub SetCaseModHistLinearDampConstant()

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



   'add linear modal history load case

      ret = SapModel.LoadCases.ModHistLinear.SetCase("LCASE1")



   'set constant damping

      ret = SapModel.LoadCases.ModHistLinear.SetDampConstant("LCASE1", 0.04)



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

[GetDampConstant](GetDampConstant_%7BModal_TH_Linear%7D.md)

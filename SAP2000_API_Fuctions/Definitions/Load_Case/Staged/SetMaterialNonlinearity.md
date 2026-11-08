# SetMaterialNonlinearity

## Syntax

SapObject.SapModel.LoadCases.StaticNonlinearStaged.SetMaterialNonlinearity

## VB6 Procedure

Function SetMaterialNonlinearity(ByVal Name As String, ByVal TimeDepMatProp As Boolean) As Long

## Parameters

Name

The name of an existing static nonlinear staged analysis case.

TimeDepMatProp

When this is True, any specified time dependent material properties are considered in the analysis.

## Remarks

This function sets the material nonlinearity options for the specified load case.

The function returns zero if the options are successfully set; otherwise it returns a nonzero value.

## VBA Example



```
Sub SetCaseStaticNonlinearStagedMaterialNonlinearity()

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



   'add static nonlinear staged load case

      ret = SapModel.LoadCases.StaticNonlinearStaged.SetCase("LCASE1")



   'set geometric nonlinearity options

      ret = SapModel.LoadCases.StaticNonlinearStaged.SetMaterialNonlinearity("LCASE1", True)



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

[GetMaterialNonlinearity](GetMaterialNonlinearity.md)

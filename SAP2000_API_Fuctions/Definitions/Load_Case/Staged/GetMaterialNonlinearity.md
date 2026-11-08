# GetMaterialNonlinearity

## Syntax

SapObject.SapModel.LoadCases.StaticNonlinearStaged.GetMaterialNonlinearity

## VB6 Procedure

Function GetMaterialNonlinearity(ByVal Name As String, ByRef TimeDepMatProp As Boolean) As Long

## Parameters

Name

The name of an existing static nonlinear staged analysis case.

TimeDepMatProp

When this is True, any specified time dependent material properties are considered in the analysis.

## Remarks

This function retrieves the material nonlinearity options for the specified load case.

The function returns zero if the options are successfully retrieved; otherwise it returns a nonzero value.

## VBA Example



```
Sub GetCaseStaticNonlinearStagedMaterialNonlinearity()

   'dimension variables

      Dim SapObject as cOAPI

      Dim SapModel As cSapModel

      Dim ret As Long

      Dim TimeDepMatProp As Boolean



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



   'get material nonlinearity options

      ret = SapModel.LoadCases.StaticNonlinearStaged.GetMaterialNonlinearity("LCASE1", TimeDepMatProp)



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

[SetMaterialNonlinearity](SetMaterialNonlinearity.md)

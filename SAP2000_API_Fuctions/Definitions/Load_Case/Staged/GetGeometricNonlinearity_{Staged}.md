# GetGeometricNonlinearity

## Syntax

SapObject.SapModel.LoadCases.StaticNonlinearStaged.GetGeometricNonlinearity

## VB6 Procedure

Function GetGeometricNonlinearity(ByVal Name As String, ByRef NLGeomType As Long) As Long

## Parameters

Name

The name of an existing static nonlinear staged analysis case.

NLGeomType

This is 0, 1 or 2, indicating the geometric nonlinearity option selected for the load case.

0 = None

1 = P-delta

2 = P-delta plus large displacements

## Remarks

This function retrieves the geometric nonlinearity option for the specified load case.

The function returns zero if the option is successfully retrieved; otherwise it returns a nonzero value.

## VBA Example



```
Sub GetCaseStaticNonlinearStagedGeometricNonlinearity()

   'dimension variables

      Dim SapObject as cOAPI

      Dim SapModel As cSapModel

      Dim ret As Long

      Dim NLGeomType As Long



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



   'get geometric nonlinearity option

      ret = SapModel.LoadCases.StaticNonlinearStaged.GetGeometricNonlinearity("LCASE1", NLGeomType)



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

[SetGeometricNonlinearity](GetGeometricNonlinearity_%7BStaged%7D.md)

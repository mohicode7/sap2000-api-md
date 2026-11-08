# SetGeometricNonlinearity

## Syntax

SapObject.SapModel.LoadCases.DirHistNonlinear.SetGeometricNonlinearity

## VB6 Procedure

Function SetGeometricNonlinearity(ByVal Name As String, ByVal NLGeomType As Long) As Long

## Parameters

Name

The name of an existing nonlinear direct integration time history load case.

NLGeomType

This is 0, 1 or 2, indicating the geometric nonlinearity option selected for the load case.

0 = None

1 = P-delta

2 = P-delta plus large displacements

## Remarks

This function sets the geometric nonlinearity option for the specified load case.

The function returns zero if the option is successfully set; otherwise it returns a nonzero value.

## VBA Example



```
Sub SetCaseDirHistNonlinearGeometricNonlinearity()

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



   'add nonlinear direct history load case

      ret = SapModel.LoadCases.DirHistNonlinear.SetCase("LCASE1")



   'set geometric nonlinearity option

      ret = SapModel.LoadCases.DirHistNonlinear.SetGeometricNonlinearity("LCASE1", 2)



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

[GetGeometricNonlinearity](GetGeometricNonlinearity_%7BDirect_TH_Nonlinear%7D.md)

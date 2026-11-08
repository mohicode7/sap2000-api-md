# SetParameters

## Syntax

SapObject.SapModel.LoadCases.Buckling.SetParameters

## VB6 Procedure

Function SetParameters(ByVal Name As String, ByVal NumBucklingModes As Long, ByVal EigenTol As Double) As Long

## Parameters

Name

The name of an existing buckling load case.

NumBucklingModes

The number of buckling modes requested.

EigenTol

The relative convergence tolerance for eigenvalues.

## Remarks

This function sets various parameters for the specified buckling load case.

The function returns zero if the parameters are successfully set; otherwise it returns a nonzero value.

## VBA Example



```
Sub SetCaseBucklingParameters()

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



   'add buckling load case

      ret = SapModel.LoadCases.Buckling.SetCase("LCASE1")



   'set parameters

      ret = SapModel.LoadCases.Buckling.SetParameters("LCASE1", 4, 1E-10)



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

[GetParameters](GetParameters_%7BBuckling%7D.md)

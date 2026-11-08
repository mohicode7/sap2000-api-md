# GetParameters

## Syntax

SapObject.SapModel.LoadCases.Buckling.GetParameters

## VB6 Procedure

Function GetParameters(ByVal Name As String, ByRef NumBucklingModes As Long, ByRef EigenTol As Double) As Long

## Parameters

Name

The name of an existing buckling load case.

NumBucklingModes

The number of buckling modes requested.

EigenTol

The relative convergence tolerance for eigenvalues.

## Remarks

This function retrieves various parameters for the specified load case.

The function returns zero if the parameters are successfully retrieved; otherwise it returns a nonzero value.

## VBA Example



```
Sub GetCaseBucklingParameters()

   'dimension variables

      Dim SapObject as cOAPI

      Dim SapModel As cSapModel

      Dim ret As Long

      Dim NumBucklingModes As Long

      Dim EigenTol As Double



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



   'get parameters

      ret = SapModel.LoadCases.Buckling.GetParameters("LCASE1", NumBucklingModes, EigenTol)



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

[SetParameters](SetParameters_%7BBuckling%7D.md)

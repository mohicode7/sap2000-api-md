# SetParameters

## Syntax

SapObject.SapModel.LoadCases.ModalEigen.SetParameters

## VB6 Procedure

Function SetParameters(ByVal Name As String, ByVal EigenShiftFreq As Double, ByVal EigenCutOff As Double, ByVal EigenTol As Double, ByVal AllowAutoFreqShift As Long) As Long

## Parameters

Name

The name of an existing modal eigen load case.

EigenShiftFreq

The eigenvalue shift frequency. [cyc/s]

EigenCutOff

The eigencutoff frequency radius. [cyc/s]

EigenTol

The relative convergence tolerance for eigenvalues.

AllowAutoFreqShift

This is either 0 or 1, indicating if automatic frequency shifting is allowed.

0 = Automatic frequency shifting is NOT allowed

1 = Automatic frequency shifting is allowed

## Remarks

This function sets various parameters for the specified modal eigen load case.

The function returns zero if the parameters are successfully set; otherwise it returns a nonzero value.

## VBA Example



```
Sub SetCaseModalEigenParameters()

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



   'add modal eigen load case

      ret = SapModel.LoadCases.ModalEigen.SetCase("LCASE1")



   'set parameters

      ret = SapModel.LoadCases.ModalEigen.SetParameters("LCASE1", 0.05, 0.0001, 1E-10, 1)



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

[GetParameters](GetParameters_%7BModal_Eigen%7D.md)

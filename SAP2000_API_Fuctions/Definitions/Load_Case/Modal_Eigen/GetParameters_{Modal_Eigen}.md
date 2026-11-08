# GetParameters

## Syntax

SapObject.SapModel.LoadCases.ModalEigen.GetParameters

## VB6 Procedure

Function GetParameters(ByVal Name As String, ByRef EigenShiftFreq As Double, ByRef EigenCutOff As Double, ByRef EigenTol As Double, ByRef AllowAutoFreqShift As Long) As Long

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

This function retrieves various parameters for the specified load case.

The function returns zero if the parameters are successfully retrieved; otherwise it returns a nonzero value.

## VBA Example



```
Sub GetCaseModalEigenParameters()

   'dimension variables

      Dim SapObject as cOAPI

      Dim SapModel As cSapModel

      Dim ret As Long

      Dim EigenShiftFreq As Double

      Dim EigenCutOff As Double

      Dim EigenTol As Double

      Dim AllowAutoFreqShift As Long



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



   'get parameters

      ret = SapModel.LoadCases.ModalEigen.GetParameters("LCASE1", EigenShiftFreq, EigenCutOff, EigenTol, AllowAutoFreqShift)



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

[SetParameters](SetParameters_%7BModal_Eigen%7D.md)

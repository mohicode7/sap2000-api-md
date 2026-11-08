# SetFreqData

## Syntax

SapObject.SapModel.LoadCases.PSD.SetFreqData

## VB6 Procedure

Function SetFreqData(ByVal Name As String, ByVal FreqFirst As Double, ByVal FreqLast As Double, ByVal FreqNumIncs As Long, ByVal FreqAddModal As Boolean, ByVal FreqAddModalDev As Boolean, ByVal FreqAddSpecified As Boolean, ByVal FreqNumModalDev As Long, ByRef FreqModalDev() As Double, ByVal FreqNumSpecified As Long, ByRef FreqSpecified() As Double) As Long

## Parameters

Name

The name of an existing power spectral density analysis case.

FreqFirst

The first frequency. [cyc/s]

FreqLast

The last frequency. [cyc/s]

FreqNumIncs

The number of frequency increments.

FreqAddModal

If this item is True, modal frequencies are added.

FreqAddModalDev

If this item is True, signed fractional deviations from modal frequencies are added.

FreqAddSpecified

If this item is True, specified frequencies are added.

ModalCase

This is the name of an existing modal load case. It specifies the modal load case on which modal frequencies and modal frequency deviations are based.

FreqNumModalDev

The number of signed fractional deviations from modal frequencies that are added. This item applies only when FreqAddModalDev = True.

FreqModalDev

This is an array that includes the added signed fractional deviations from modal frequencies. This item applies only when FreqAddModalDev = True.

FreqNumSpecified

The number of specified frequencies that are added. This item applies only when FreqAddSpecified = True.

FreqSpecified

This is an array that includes the added specified frequencies. This item applies only when FreqAddModalDev = True.

## Remarks

This function sets the frequency data for the specified load case.

The function returns zero if the data is successfully set; otherwise it returns a nonzero value.

## VBA Example



```
Sub SetCasePSDFreqData()

   'dimension variables

      Dim SapObject as cOAPI

      Dim SapModel As cSapModel

      Dim ret As Long

      Dim MyFreqModalDev() As Double

      Dim MyFreqSpecified() As Double



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



   'add power spectral density load case

      ret = SapModel.LoadCases.PSD.SetCase("LCASE1")



   'set frequency data

      ReDim MyFreqModalDev(1)

      ReDim MyFreqSpecified(1)

      MyFreqModalDev(0) = -0.1

      MyFreqModalDev(1) = 0.2

      MyFreqSpecified(0) = 1.2

      MyFreqSpecified(1) = 11.4

      ret = SapModel.LoadCases.PSD.SetFreqData("LCASE1", .6, 20.6, 10, True, True, True, "MODAL", 2, MyFreqModalDev, 2, MyFreqSpecified)



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

[GetFreqData](GetFreqData_%7BPSD%7D.md)

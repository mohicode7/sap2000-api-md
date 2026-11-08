# SetModalComb (Note:  Newer function available)

## Syntax

SapObject.SapModel.LoadCases.ResponseSpectrum.SetModalComb

## VB6 Procedure

Function SetModalComb(ByVal Name As String, ByVal MyType As Long, Optional ByVal F1 As Double = 1, Optional ByVal F2 As Double = 0, Optional ByVal td As Double = 60) As Long

## Parameters

Name

The name of an existing response spectrum load case.

MyType

This is 1, 2, 3, 4, 5 or 6, indicating the modal combination option.

1 = CQC

2 = SRSS

3 = ABS

4 = GMC

5 = 10 percent

6 = Double sum

F1

This item applies only when MyType = 4. It is the GMC f1 factor. [cyc/s]

F2

This item applies only when MyType = 4. It is the GMC f2 factor. [cyc/s]

td

This item applies only when MyType = 6. It is the factor td. [s]

## Remarks

This function sets the modal combination option for the specified load case.

The function returns zero if the option is successfully set; otherwise it returns a nonzero value.

## VBA Example



```
Sub SetCaseResponseSpectrumModalComb()

   'dimension variables

      Dim SapObject As Sap2000v15.SapObject

      Dim SapModel As cSapModel

      Dim ret As Long



   'create Sap2000 object

      Set SapObject = New Sap2000v15.SapObject



   'start Sap2000 application

      SapObject.ApplicationStart



   'create SapModel object

      Set SapModel = SapObject.SapModel



   'initialize model

      ret = SapModel.InitializeNewModel



   'create model from template

      ret = SapModel.File.New2DFrame(PortalFrame, 2, 144, 2, 288)



   'add response spectrum load case

      ret = SapModel.LoadCases.ResponseSpectrum.SetCase("LCASE1")



   'set modal combination option

      ret = SapModel.LoadCases.ResponseSpectrum.SetModalComb("LCASE1", 4, 2, 0.5)



   'close Sap2000

      SapObject.ApplicationExit False

      Set SapModel = Nothing

      Set SapObject = Nothing

End Sub
```



## Release Notes

Initial release in version 11.02.

Changed nomenclature from Load Cases, Analysis Cases and Response Combinations to Load Patterns, Load Cases and Load Combinations, respectively, in version 12.00.

The function is obsolete and has been superseded by [SetModalComb\_1](../Definitions/Load_Case/Response_Spectrum/SetModalComb_1.md) as of version 14.00. This function is maintained for backward compatibility. New function added

## See Also

[GetModalComb](GetModalComb.md)

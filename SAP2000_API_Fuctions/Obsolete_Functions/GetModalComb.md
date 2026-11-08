# GetModalComb (Note:  Newer function available)

## Syntax

SapObject.SapModel.LoadCases.ResponseSpectrum.GetModalComb

## VB6 Procedure

Function GetModalComb(ByVal Name As String, ByRef MyType As Long, ByRef F1 As Double, ByRef F2 As Double, ByRef td As Double) As Long

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

This function retrieves the modal combination option assigned to the specified load case.

The function returns zero if the option is successfully retrieved; otherwise it returns a nonzero value.

## VBA Example



```
Sub GetCaseResponseSpectrumModalComb()

   'dimension variables

      Dim SapObject As Sap2000v15.SapObject

      Dim SapModel As cSapModel

      Dim ret As Long

      Dim MyType As Long

      Dim F1 As Double

      Dim F2 As Double

      Dim td As Double



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



   'get modal combination option

      ret = SapModel.LoadCases.ResponseSpectrum.GetModalComb("LCASE1", MyType, F1, F2, td)



   'close Sap2000

      SapObject.ApplicationExit False

      Set SapModel = Nothing

      Set SapObject = Nothing

End Sub
```



## Release Notes

Initial release in version 11.02.

Changed nomenclature from Load Cases, Analysis Cases and Response Combinations to Load Patterns, Load Cases and Load Combinations, respectively, in version 12.00.

The function is obsolete and has been superseded by [GetModalComb\_1](../Definitions/Load_Case/Response_Spectrum/GetModalComb_1.md) as of version 14.00. This function is maintained for backward compatibility. New function added

## See Also

[SetModalComb](GetModalComb.md)

# GetDirComb

## Syntax

SapObject.SapModel.LoadCases.ResponseSpectrum.GetDirComb

## VB6 Procedure

Function GetDirComb(ByVal Name As String, ByRef MyType As Long, ByRef SF As Double) As Long

## Parameters

Name

The name of an existing response spectrum load case.

MyType

This is 1, 2, or 3, indicating the directional combination option.

1 = SRSS

2 = ABS  
3 = CQC3

SF

This item applies only when MyType = 2. It is the ABS scale factor.

## Remarks

This function retrieves the directional combination option assigned to the specified load case.

The function returns zero if the option is successfully retrieved; otherwise it returns a nonzero value.

## VBA Example



```
Sub GetCaseResponseSpectrumDirComb()

   'dimension variables

      Dim SapObject as cOAPI

      Dim SapModel As cSapModel

      Dim ret As Long

      Dim MyType As Long

      Dim SF As Double



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



   'add response spectrum load case

      ret = SapModel.LoadCases.ResponseSpectrum.SetCase("LCASE1")



   'get directional combination option

      ret = SapModel.LoadCases.ResponseSpectrum.GetDirComb("LCASE1", MyType, SF)



   'close Sap2000

      SapObject.ApplicationExit False

      Set SapModel = Nothing

      Set SapObject = Nothing

End Sub
```



## Release Notes

Initial release in version 11.02.

Changed nomenclature from Load Cases, Analysis Cases and Response Combinations to Load Patterns, Load Cases and Load Combinations, respectively, in version 12.00.

Removed Modified SRSS (Chinese) directional combination option in version 14.00.

Added CQC3 directional combination option in version 14.20.

## See Also

[SetDirComb](SetDirComb.md)

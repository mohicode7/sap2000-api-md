# GetModalComb\_1

## Syntax

SapObject.SapModel.LoadCases.ResponseSpectrum.GetModalComb\_1

## VB6 Procedure

Function GetModalComb\_1(ByVal Name As String, ByRef MyType As Long, ByRef F1 As Double, ByRef F2 As Double, ByRef PeriodicRigidCombType As Long, ByRef td As Double) As Long

## Parameters

Name

The name of an existing response spectrum load case.

MyType

This is 1, 2, 3, 4, 5 or 6, indicating the modal combination option.

1 = CQC

2 = SRSS

3 = Absolute

4 = GMC

5 = NRC 10 percent

6 = Double sum

F1

The GMC f1 factor. This item does not apply when MyType = 3. [cyc/s]

F2

The GMC f2 factor. This item does not apply when MyType = 3. [cyc/s]

PeriodicRigidCombType

This is 1 or 2, indicating the periodic plus rigid modal combination option.

1 = SRSS

2 = Absolute

td

This item applies only when MyType = 6. It is the factor td. [s]

## Remarks

This function retrieves the modal combination option assigned to the specified load case.

The function returns zero if the option is successfully retrieved; otherwise it returns a nonzero value.

## VBA Example



```
Sub GetCaseResponseSpectrumModalComb_1()

   'dimension variables

      Dim SapObject as cOAPI

      Dim SapModel As cSapModel

      Dim ret As Long

      Dim MyType As Long

      Dim F1 As Double

      Dim F2 As Double

      Dim PeriodicRigidCombType As Long

      Dim td As Double



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



   'get modal combination option

      ret = SapModel.LoadCases.ResponseSpectrum.GetModalComb_1("LCASE1", MyType,F1, F2, PeriodicRigidCombType, td)



   'close Sap2000

      SapObject.ApplicationExit False

      Set SapModel = Nothing

      Set SapObject = Nothing

   End Sub
```



## Release Notes

Initial release in version 14.00.

This function supersedes [GetModalComb](../../../Obsolete_Functions/GetModalComb.md).

## See Also

[SetModalComb\_1](SetModalComb_1.md)

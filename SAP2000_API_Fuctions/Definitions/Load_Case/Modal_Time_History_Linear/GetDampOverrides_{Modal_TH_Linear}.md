# GetDampOverrides

## Syntax

SapObject.SapModel.LoadCases.ModHistLinear.GetDampInterpolated

## VB6 Procedure

Function GetDampInterpolated(ByVal Name As String, ByRef NumberItems As Long, ByRef Mode() As Long, ByRef Damp() As Double) As Long

## Parameters

Name

The name of an existing linear modal history analysis case.

NumberItems

The number of Mode and Damp pairs.

Mode

This is an array that includes a mode number.

Damp

This is an array that includes the damping for the specified mode (0 <= Damp < 1).

## Remarks

This function retrieves the modal damping overrides assigned to the specified load case.

The function returns zero if the overrides are successfully retrieved; otherwise it returns a nonzero value.

## VBA Example



```
Sub GetCaseModHistLinearDampOverrides()

   'dimension variables

      Dim SapObject as cOAPI

      Dim SapModel As cSapModel

      Dim ret As Long

      Dim MyMode() As Long

      Dim MyDamp() As Double

      Dim NumberItems As Long

      Dim Mode() As Long

      Dim Damp() As Double



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



   'add linear modal history load case

      ret = SapModel.LoadCases.ModHistLinear.SetCase("LCASE1")



   'set constant damping

      ret = SapModel.LoadCases.ModHistLinear.SetDampConstant("LCASE1", 0.04)



   'set modal damping overrides

      ReDim MyMode(1)

      ReDim MyDamp(1)

      MyMode(0) = 1

      MyDamp(0) = 0.02

      MyMode(1) = 2

      MyDamp(1) = 0.03

      ret = SapModel.LoadCases.ModHistLinear.SetDampOverrides("LCASE1", 2, MyMode, MyDamp)



   'get modal damping overrides

      ret = SapModel.LoadCases.ModHistLinear.GetDampOverrides("LCASE1", NumberItems, Mode, Damp)



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

[SetDampOverrides](SetDampOverrides_%7BModal_TH_Linear%7D.md)

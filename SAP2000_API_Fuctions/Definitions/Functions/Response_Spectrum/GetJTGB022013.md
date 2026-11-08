# GetJTGB022013

## Syntax

SapObject.SapModel.Func.FuncRS.GetJTGB022013

## VB6 Procedure

Function GetJTGB022013(ByVal Name As String, ByRef Direction As Integer, ByRef PeakAccel As Double, ByRef Tg As Double, ByRef Ci As Double, ByRef Cs As Double, ByRef DampRatio As Double) As Long

## Parameters

Name

The name of a JTG B02-2013 response spectrum function.

Direction

This is 1, 2 or 3, indicating the response spectrum direction.

1 = Horizontal

2 = Vertical-Rock

3 = Vertical-Soil

PeakAccel

The peak acceleration, A.

Tg

The characteristic ground period, Tg > 0.1. [s]

Ci

The importance coefficient.

Cs

The site soil coefficient.

DampRatio

The damping ratio for the function, 0 <= DampRatio < 1.

## Remarks

This function retrieves the definition of a JTG B02-2013 response spectrum function.

The function returns zero if the function definition is successfully retrieved; otherwise it returns a nonzero value.

## VBA Example



```
Sub GetRSFuncJTGB022013()

   'dimension variables

      Dim SapObject as cOAPI

      Dim SapModel As cSapModel

      Dim ret As Long

      Dim Direction As Long

      Dim PeakAccel As Double

      Dim Tg As Double

      Dim Ci As Double

      Dim Cs As Double

      Dim DampRatio As Double



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



   'add JTG B02-2013 RS function

      ret = SapModel.Func.FuncRS.SetJTGB022013("RS-1", 1, 0.18, 0.36, 0.4, 1, 0.04)
   'get JTG B02-2013 RS function

      ret = SapModel.Func.FuncRS.GetJTGB022013("RS-1", Direction, PeakAccel, Tg, Ci, Cs, DampRatio)

  

   'close Sap2000

      SapObject.ApplicationExit False

      Set SapModel = Nothing

      Set SapObject = Nothing

End Sub
```



## Release Notes

Initial release in v18.2.0.

## See Also

[SetJTBG022013](SetJTBG022013.md)

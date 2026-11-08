# GetCJJ1662011 {RS}

## Syntax

SapObject.SapModel.Func.FuncRS.GetCJJ1662011

## VB6 Procedure

Function GetCJJ1662011(ByVal Name As String, ByRef Direction As Integer, ByRef PeakAccel As Double, ByRef Tg As Double, ByRef DampRatio As Double) As Long

## Parameters

Name

The name of a CJJ 166-2011 response spectrum function.

Direction

This is 1 or 2, indicating the response spectrum direction.

1 = Horizontal

2 = Vertical

PeakAccel

The peak acceleration, A.

Tg

The characteristic ground period , Tg > 0.1. [s]

DampRatio

The damping ratio for the function, 0 <= DampRatio < 1.

## Remarks

This function retrieves the definition of a CJJ 166-2011 response spectrum function.

The function returns zero if the function is successfully retrieved; otherwise it returns a nonzero value.

## VBA Example



```
Sub GetRSFuncCJJ1662011()

   'dimension variables

      Dim SapObject as cOAPI

      Dim SapModel As cSapModel

      Dim ret As Long
      Dim Direction As Long
      Dim PeakAccel As Double
      Dim Tg As Double
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



   'add CJJ 166-2011 RS function

      ret = SapModel.Func.FuncRS.SetCJJ1662011("RS-1", 0.35, 0.25, 0.04)
   'get CJJ 166-2011 RS function
      ret = SapModel.Func.FuncRS.GetCJJ1662011("RS-1", Direction, PeakAccel, Tg, DampRatio)
   'close Sap2000

      SapObject.ApplicationExit False

      Set SapModel = Nothing

      Set SapObject = Nothing

End Sub
```



## Release Notes

Initial release in version 18.2.0.

## See Also

[SetCJJ1662011](SetCJJ1662011.md)

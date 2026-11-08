# GetChinese2010

## Syntax

SapObject.SapModel.Func.FuncRS.GetChinese2010

## VB6 Procedure

Function GetChinese2010(ByVal Name As String, ByRef JGJ32010AlphaMax As Double, ByRef JGJ32010SI As Long, ByRef JGJ32010Tg As Double, ByRef JGJ32010PTDF As Double, ByRef DampRatio As Double) As Long

## Parameters

Name

The name of a Chinese 2010 response spectrum function.

JGJ32010AlphaMax

The maximum influence factor.

JGJ32010SI

This is 1, 2, 3, 4, 5 or 6, indicating the seismic intensity.

1 = 6 (0.05g)

2 = 7 (0.10g)

3 = 7 (0.15g)

4 = 8 (0.20g)

5 = 8 (0.30g)

6 = 9 (0.40g)

JGJ32010Tg

The characteristic ground period, Tg > 0.1. [s]

JGJ32010PTDF

The period time discount factor.

DampRatio

The damping ratio for the function, 0 <= DampRatio < 1.

## Remarks

This function retrieves the definition of a Chinese 2010 response spectrum function.

The function returns zero if the function definition is successfully retrieved; otherwise it returns a nonzero value.

## VBA Example



```
Sub GetRSFuncChinese2010()

   'dimension variables

      Dim SapObject as cOAPI

      Dim SapModel As cSapModel

      Dim ret As Long

      Dim JGJ32010AlphaMax As Double

      Dim JGJ32010SI As Long

      Dim JGJ32010Tg As Double

      Dim JGJ32010PTDF As Double

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



   'add Chinese2010 RS function

      ret = SapModel.Func.FuncRS.SetChinese2010("RS-1", 0.18, 5, 0.36, 1, 0.04)



   'get Chinese2010 RS function

      ret = SapModel.Func.FuncRS.GetChinese2010("RS-1", JGJ32010AlphaMax, JGJ32010SI, JGJ32010Tg, JGJ32010PTDF, DampRatio)



   'close Sap2000

      SapObject.ApplicationExit False

      Set SapModel = Nothing

      Set SapObject = Nothing

End Sub
```



## Release Notes

Initial release in version 15.0.2.

## See Also

[SetChinese2010](SetChinese2010_%7BRS%7D.md)

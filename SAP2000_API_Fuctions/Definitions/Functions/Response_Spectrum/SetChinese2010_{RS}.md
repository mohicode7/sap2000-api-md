# SetChinese2010

## Syntax

SapObject.SapModel.Func.FuncRS.SetChinese2010

## VB6 Procedure

Function SetChinese2010(ByVal Name As String, ByVal JGJ32010AlphaMax As Double, ByVal JGJ32010SI As Long, ByVal JGJ32010Tg As Double, ByVal JGJ32010PTDF As Double, ByVal DampRatio As Double) As Long

## Parameters

Name

The name of an existing or new function. If this is an existing function, that function is modified; otherwise, a new function is added.

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

This function defines a Chinese 2010 response spectrum function.

The function returns zero if the function is successfully defined; otherwise it returns a nonzero value.

## VBA Example



```
Sub SetRSFuncChinese2010()

   'dimension variables

      Dim SapObject as cOAPI

      Dim SapModel As cSapModel

      Dim ret As Long



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



   'close Sap2000

      SapObject.ApplicationExit False

      Set SapModel = Nothing

      Set SapObject = Nothing

End Sub
```



## Release Notes

Initial release in version 15.0.2.

## See Also

[GetChinese2010](GetChinese2010_%7BRS%7D.md)

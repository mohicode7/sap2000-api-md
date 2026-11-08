# GetChinese2002

## Syntax

SapObject.SapModel.Func.FuncRS.GetChinese2002

## VB6 Procedure

Function GetChinese2002(ByVal Name As String, ByRef JGJ32002AlphaMax As Double, ByRef JGJ32002SI As Long, ByRef JGJ32002Tg As Double, ByRef JGJ32002PTDF As Double, ByRef DampRatio As Double) As Long

## Parameters

Name

The name of a Chinese 2002 response spectrum function.

JGJ32002AlphaMax

The maximum influence factor.

JGJ32002SI

This is 1, 2, 3, 4, 5 or 6, indicating the seismic intensity.

1 = 6 (0.05g)

2 = 7 (0.10g)

3 = 7 (0.15g)

4 = 8 (0.20g)

5 = 8 (0.30g)

6 = 9 (0.40g)

JGJ32002Tg

The characteristic ground period, Tg > 0.1. [s]

JGJ32002PTDF

The period time discount factor.

DampRatio

The damping ratio for the function, 0 <= DampRatio < 1.

## Remarks

This function retrieves the definition of a Chinese 2002 response spectrum function.

The function returns zero if the function definition is successfully retrieved; otherwise it returns a nonzero value.

## VBA Example



```
Sub GetRSFuncChinese2002()

   'dimension variables

      Dim SapObject As Sap2000v15.SapObject

      Dim SapModel As cSapModel

      Dim ret As Long

      Dim JGJ32002AlphaMax As Double

      Dim JGJ32002SI As Long

      Dim JGJ32002Tg As Double

      Dim JGJ32002PTDF As Double

      Dim DampRatio As Double



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



   'add Chinese2002 RS function

      ret = SapModel.Func.FuncRS.SetChinese2002("RS-1", 0.18, 5, 0.36, 1, 0.04)



   'get Chinese2002 RS function

      ret = SapModel.Func.FuncRS.GetChinese2002("RS-1", JGJ32002AlphaMax, JGJ32002SI, JGJ32002Tg, JGJ32002PTDF, DampRatio)



   'close Sap2000

      SapObject.ApplicationExit False

      Set SapModel = Nothing

      Set SapObject = Nothing

End Sub
```



## Release Notes

Initial release in version 11.02.

## See Also

[SetChinese2002](SetChinese2002_%7BRS%7D.md)

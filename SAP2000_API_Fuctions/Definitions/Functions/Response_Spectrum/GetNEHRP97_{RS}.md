# GetNEHRP97

## Syntax

SapObject.SapModel.Func.FuncRS.GetNEHRP97

## VB6 Procedure

Function GetNEHRP97(ByVal Name As String, ByRef NEHRP97SS As Double, ByRef NEHRP97S1 As Double, ByRef DampRatio As Double) As Long

## Parameters

Name

The name of a NEHRP97 response spectrum function.

NEHRP97SS

The design spectral acceleration at short periods, Sds.

NEHRP97S1

The design spectral acceleration at a one second period, Sd1.

DampRatio

The damping ratio for the function, 0 <= DampRatio < 1.

## Remarks

This function retrieves the definition of a NEHRP97 response spectrum function.

The function returns zero if the function definition is successfully retrieved; otherwise it returns a nonzero value.

## VBA Example



```
Sub GetRSFuncNEHRP97()

   'dimension variables

      Dim SapObject as cOAPI

      Dim SapModel As cSapModel

      Dim ret As Long

      Dim NEHRP97SS As Double

      Dim NEHRP97S1 As Double

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



   'add NEHRP97 RS function

      ret = SapModel.Func.FuncRS.SetNEHRP97("RS-1", 1.2, 0.3, 0.04)



   'get NEHRP97 RS function

      ret = SapModel.Func.FuncRS.GetNEHRP97("RS-1", NEHRP97SS, NEHRP97S1, DampRatio)



   'close Sap2000

      SapObject.ApplicationExit False

      Set SapModel = Nothing

      Set SapObject = Nothing

End Sub
```



## Release Notes

Initial release in version 11.02.

## See Also

[SetNEHRP97](GetNEHRP97_%7BRS%7D.md)

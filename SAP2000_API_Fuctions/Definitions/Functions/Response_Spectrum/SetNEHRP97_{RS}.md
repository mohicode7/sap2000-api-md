# SetNEHRP97

## Syntax

SapObject.SapModel.Func.FuncRS.SetNEHRP97

## VB6 Procedure

Function SetNEHRP97(ByVal Name As String, ByVal NEHRP97SS As Double, ByVal NEHRP97S1 As Double, ByVal DampRatio As Double) As Long

## Parameters

Name

The name of an existing or new function. If this is an existing function then that function is modified; otherwise, a new function is added.

NEHRP97SS

The design spectral acceleration at short periods, Sds.

NEHRP97S1

The design spectral acceleration at a one second period, Sd1.

DampRatio

The damping ratio for the function, 0 <= DampRatio < 1.

## Remarks

This function defines a NEHRP97 response spectrum function.

The function returns zero if the function is successfully defined; otherwise it returns a nonzero value.

## VBA Example



```
Sub SetRSFuncNEHRP97()

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



   'add NEHRP97 RS function

      ret = SapModel.Func.FuncRS.SetNEHRP97("RS-1", 1.2, 0.3, 0.04)



   'close Sap2000

      SapObject.ApplicationExit False

      Set SapModel = Nothing

      Set SapObject = Nothing

End Sub
```



## Release Notes

Initial release in version 11.02.

## See Also

[GetNEHRP97](GetNEHRP97_%7BRS%7D.md)

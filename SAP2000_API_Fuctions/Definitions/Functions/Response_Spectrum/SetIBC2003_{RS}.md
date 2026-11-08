# SetIBC2003

## Syntax

SapObject.SapModel.Func.FuncRS.SetIBC2003

## VB6 Procedure

Function SetIBC2003(ByVal Name As String, ByVal IBC2003SS As Double, ByVal IBC2003S1 As Double, ByVal DampRatio As Double) As Long

## Parameters

Name

The name of an existing or new function. If this is an existing function then that function is modified, otherwise, a new function is added.

IBC2003SS

The design spectral acceleration at short periods, Sds.

IBC2003S1

The design spectral acceleration at a one second period, Sd1.

DampRatio

The damping ratio for the function, 0 <= DampRatio < 1.

## Remarks

This function defines a IBC2003 response spectrum function.

The function returns zero if the function is successfully defined; otherwise it returns a nonzero value.

## VBA Example



```
Sub SetRSFuncIBC2003()

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



   'add IBC2003 RS function

      ret = SapModel.Func.FuncRS.SetIBC2003("RS-1", 1.2, 0.3, 0.04)



   'close Sap2000

      SapObject.ApplicationExit False

      Set SapModel = Nothing

      Set SapObject = Nothing

End Sub
```



## Release Notes

Initial release in version 11.02.

## See Also

[GetIBC2003](GetIBC2003_%7BRS%7D.md)

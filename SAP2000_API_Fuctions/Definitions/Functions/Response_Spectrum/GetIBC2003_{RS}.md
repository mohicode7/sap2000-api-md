# GetIBC2003

## Syntax

SapObject.SapModel.Func.FuncRS.GetIBC2003

## VB6 Procedure

Function GetIBC2003(ByVal Name As String, ByRef IBC2003SS As Double, ByRef IBC2003S1 As Double, ByRef DampRatio As Double) As Long

## Parameters

Name

The name of a IBC2003 response spectrum function.

IBC2003SS

The design spectral acceleration at short periods, Sds.

IBC2003S1

The design spectral acceleration at a one second period, Sd1.

DampRatio

The damping ratio for the function, 0 <= DampRatio < 1.

## Remarks

This function retrieves the definition of a IBC2003 response spectrum function.

The function returns zero if the function definition is successfully retrieved; otherwise it returns a nonzero value.

## VBA Example



```
Sub GetRSFuncIBC2003()

   'dimension variables

      Dim SapObject as cOAPI

      Dim SapModel As cSapModel

      Dim ret As Long

      Dim IBC2003SS As Double

      Dim IBC2003S1 As Double

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



   'add IBC2003 RS function

      ret = SapModel.Func.FuncRS.SetIBC2003("RS-1", 1.2, 0.3, 0.04)



   'get IBC2003 RS function

      ret = SapModel.Func.FuncRS.GetIBC2003("RS-1", IBC2003SS, IBC2003S1, DampRatio)



   'close Sap2000

      SapObject.ApplicationExit False

      Set SapModel = Nothing

      Set SapObject = Nothing

End Sub
```



## Release Notes

Initial release in version 11.02.

## See Also

[SetIBC2003](SetIBC2003_%7BRS%7D.md)

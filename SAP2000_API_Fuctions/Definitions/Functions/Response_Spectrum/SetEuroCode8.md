# SetEuroCode8

## Syntax

SapObject.SapModel.Func.FuncRS.SetEuroCode8

## VB6 Procedure

Function SetEuroCode8(ByVal Name As String, ByVal EuroCode8AG As Double, ByVal EuroCode8S As Long, ByVal EuroCode8n As Double, ByVal DampRatio As Double) As Long

## Parameters

Name

The name of an existing or new function. If this is an existing function, that function is modified; otherwise, a new function is added.

EuroCode8AG

The design ground acceleration, Ag.

EuroCode8S

This is 1, 2 or 3, indicating the subsoil class.

1 = A

2 = B

3 = C

EuroCode8n

The damping correction factor, n, where n >= 0.7.

DampRatio

The damping ratio for the function, 0 <= DampRatio < 1.

## Remarks

This function defines a EuroCode8 response spectrum function.

The function returns zero if the function is successfully defined; otherwise it returns a nonzero value.

## VBA Example



```
Sub SetRSFuncEuroCode8()

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



   'add EuroCode8 RS function

      ret = SapModel.Func.FuncRS.SetEuroCode8("RS-1", 0.3, 3, 1.2, 0.04)



   'close Sap2000

      SapObject.ApplicationExit False

      Set SapModel = Nothing

      Set SapObject = Nothing

End Sub
```



## Release Notes

Initial release in version 11.02.

## See Also

[GetEuroCode8](GetEuroCode8.md)

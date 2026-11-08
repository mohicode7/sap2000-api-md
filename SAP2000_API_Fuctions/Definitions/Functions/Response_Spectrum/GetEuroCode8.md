# GetEuroCode8

## Syntax

SapObject.SapModel.Func.FuncRS.GetEuroCode8

## VB6 Procedure

Function GetEuroCode8(ByVal Name As String, ByRef EuroCode8AG As Double, ByRef EuroCode8S As Long, ByRef EuroCode8n As Double, ByRef DampRatio As Double) As Long

## Parameters

Name

The name of a EuroCode8 response spectrum function.

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

This function retrieves the definition of a EuroCode8 response spectrum function.

The function returns zero if the function definition is successfully retrieved; otherwise it returns a nonzero value.

## VBA Example



```
Sub GetRSFuncEuroCode8()

   'dimension variables

      Dim SapObject as cOAPI

      Dim SapModel As cSapModel

      Dim ret As Long

      Dim EuroCode8AG As Double

      Dim EuroCode8S As Long

      Dim EuroCode8n As Double

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



   'add EuroCode8 RS function

      ret = SapModel.Func.FuncRS.SetEuroCode8("RS-1", 0.3, 3, 1.2, 0.04)



   'get EuroCode8 RS function

      ret = SapModel.Func.FuncRS.GetEuroCode8("RS-1", EuroCode8AG, EuroCode8S, EuroCode8n, DampRatio)



   'close Sap2000

      SapObject.ApplicationExit False

      Set SapModel = Nothing

      Set SapObject = Nothing

End Sub
```



## Release Notes

Initial release in version 11.02.

## See Also

[SetEuroCode8](SetEuroCode8.md)

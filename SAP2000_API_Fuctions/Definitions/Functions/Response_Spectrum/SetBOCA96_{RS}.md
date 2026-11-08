# SetBOCA96

## Syntax

SapObject.SapModel.Func.FuncRS.SetBOCA96

## VB6 Procedure

Function SetBOCA96(ByVal Name As String, ByVal BOCA96Aa As Double, ByVal BOCA96Av As Double, ByVal BOCA96S As Double, ByVal BOCA96R As Double, ByVal DampRatio As Double) As Long

## Parameters

Name

The name of an existing or new function. If this is an existing function, that function is modified; otherwise, a new function is added.

BOCA96Aa

The effective peak acceleration, Aa.

BOCA96Av

The effective peak velocity, Av.

BOCA96S

The soil profile coefficient, S.

BOCA96R

The response modification factor, R.

DampRatio

The damping ratio for the function, 0 <= DampRatio < 1.

## Remarks

This function defines a BOCA96 response spectrum function.

The function returns zero if the function is successfully defined; otherwise it returns a nonzero value.

## VBA Example



```
Sub SetRSFuncBOCA96()

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



   'add BOCA96 RS function

      ret = SapModel.Func.FuncRS.SetBOCA96("RS-1", 0.3, 0.3, 1.2, 1, 0.04)



   'close Sap2000

      SapObject.ApplicationExit False

      Set SapModel = Nothing

      Set SapObject = Nothing

End Sub
```



## Release Notes

Initial release in version 11.02.

## See Also

[GetBOCA96](GetBOCA96_%7BRS%7D.md)

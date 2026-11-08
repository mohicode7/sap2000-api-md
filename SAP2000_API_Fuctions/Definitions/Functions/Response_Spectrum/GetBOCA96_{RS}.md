# GetBOCA96

## Syntax

SapObject.SapModel.Func.FuncRS.GetBOCA96

## VB6 Procedure

Function GetBOCA96(ByVal Name As String, ByRef BOCA96Aa As Double, ByRef BOCA96Av As Double, ByRef BOCA96S As Double, ByRef BOCA96R As Double, ByRef DampRatio As Double) As Long

## Parameters

Name

The name of a BOCA96 response spectrum function.

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

This function retrieves the definition of a BOCA96 response spectrum function.

The function returns zero if the function definition is successfully retrieved; otherwise it returns a nonzero value.

## VBA Example



```
Sub GetRSFuncBOCA96()

   'dimension variables

      Dim SapObject as cOAPI

      Dim SapModel As cSapModel

      Dim ret As Long

      Dim BOCA96Aa As Double

      Dim BOCA96Av As Double

      Dim BOCA96S As Double

      Dim BOCA96R As Double

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



   'add BOCA96 RS function

      ret = SapModel.Func.FuncRS.SetBOCA96("RS-1", 0.3, 0.3, 1.2, 1, 0.04)



   'get BOCA96 RS function

      ret = SapModel.Func.FuncRS.GetBOCA96("RS-1", BOCA96Aa, BOCA96Av, BOCA96S, BOCA96R, DampRatio)



   'close Sap2000

      SapObject.ApplicationExit False

      Set SapModel = Nothing

      Set SapObject = Nothing

End Sub
```



## Release Notes

Initial release in version 11.02.

## See Also

[SetBOCA96](SetBOCA96_%7BRS%7D.md)

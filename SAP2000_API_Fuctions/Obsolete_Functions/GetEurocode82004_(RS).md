# GetEurocode82004  (Note:  Newer function available)

## Syntax

SapObject.SapModel.Func.FuncRS.GetEurocode82004

## VB6 Procedure

Function GetEurocode82004(ByVal Name As String, ByRef EURO2004GroundType As Long, ByRef EURO2004SpectrumType As Long, ByRef EURO2004ag As Double, ByRef EURO2004Beta As Double, ByRef EURO2004q As Double, ByRef DampRatio As Double) As Long

## Parameters

Name

The name of a Eurocode 8 2004 response spectrum function.

EURO2004GroundType

This is 1, 2, 3, 4 or 5, indicating the ground type.

1 = A

2 = B

3 = C

4 = D

5 = E

EURO2004SpectrumType

This is 1 or 2, indicating the spectrum type.

1 = Type 1

2 = Type 2

EURO2004ag

The design ground acceleration in g, ag.

EURO2004Beta

The lower bound factor, Beta.

EURO2004q

The behavior factor, q.

DampRatio

The damping ratio for the function, 0 <= DampRatio < 1.

## Remarks

This function retrieves the definition of a Eurocode 8 2004 response spectrum function.

The function returns zero if the function definition is successfully retrieved; otherwise it returns a nonzero value.

## VBA Example



```
Sub GetRSFuncEurocode82004()

   'dimension variables

      Dim SapObject As Sap2000v15.SapObject

      Dim SapModel As cSapModel

      Dim ret As Long

      Dim EURO2004GroundType As Long

      Dim EURO2004SpectrumType As Long

      Dim EURO2004ag As Double

      Dim EURO2004Beta As Double

      Dim EURO2004q As Double

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



   'add Eurocode 8 2004 RS function

      ret = SapModel.Func.FuncRS.SetEurocode82004("RS-1", 2, 1, 0.4, 0.2, 2, 0.04)



   'get Eurocode 8 2004 RS function

      ret = SapModel.Func.FuncRS.GetEurocode82004("RS-1", EURO2004GroundType, EURO2004SpectrumType, EURO2004ag, EURO2004Beta, EURO2004q, DampRatio)



   'close Sap2000

      SapObject.ApplicationExit False

      Set SapModel = Nothing

      Set SapObject = Nothing

   End Sub
```



## Release Notes

Initial release in version 14.00.

The function is obsolete and has been superseded by [GetEurocode82004\_1](../Definitions/Functions/Response_Spectrum/GetEurocode82004_1_%7BRS%7D.md) as of version 14.1.0. This function is maintained for backward compatibility. New function added.

## See Also

[SetEurocode82004](SetEurocode82004_%7BRS%7D.md)

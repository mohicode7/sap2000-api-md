# SetEurocode82004  (Note:  Newer function available)

## Syntax

SapObject.SapModel.Func.FuncRS.SetEurocode82004

## VB6 Procedure

Function SetEurocode82004(ByVal Name As String, ByVal EURO2004GroundType As Long, ByVal EURO2004SpectrumType As Long, ByVal EURO2004ag As Double, ByVal EURO2004Beta As Double, ByVal EURO2004q As Double, ByVal DampRatio As Double) As Long

## Parameters

Name

The name of an existing or new function. If this is an existing function,n that function is modified; otherwise, a new function is added.

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

This function defines a Eurocode 8 2004 response spectrum function.

The function returns zero if the function is successfully defined; otherwise it returns a nonzero value.

## VBA Example



```
Sub SetRSFuncEurocode82004()

   'dimension variables

      Dim SapObject As Sap2000v15.SapObject

      Dim SapModel As cSapModel

      Dim ret As Long



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



   'close Sap2000

      SapObject.ApplicationExit False

      Set SapModel = Nothing

      Set SapObject = Nothing

End Sub
```



## Release Notes

Initial release in version 14.00.

The function is obsolete and has been superseded by [SetEurocode82004\_1](../Definitions/Functions/Response_Spectrum/SetEurocode82004_1_%7BRS%7D.md) as of version 14.1.0. This function is maintained for backward compatibility. New function added.

## See Also

[GetEurocode82004](GetEurocode82004_%28RS%29.md)

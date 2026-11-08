# GetEurocode82004\_1

## Syntax

SapObject.SapModel.Func.FuncRS.GetEurocode82004\_1

## VB6 Procedure

Function GetEurocode82004\_1(ByVal Name As String, ByRef EURO2004Country As Long, ByRef EURO2004Direction As Long, ByRef EURO2004SpectrumType As Long, ByRef EURO2004GroundType As Long, ByRef EURO2004ag As Double, ByRef EURO2004S As Double, ByRef EURO2004AvgoverAg As Double, ByRef EURO2004Tb As Double, ByRef EURO2004Tc As Double, ByRef EURO2004Td As Double, ByRef EURO2004Beta As Double, ByRef EURO2004q As Double, ByRef DampRatio As Double) As Long

## Parameters

Name

The name of a Eurocode 8 2004 response spectrum function.

EURO2004Country

This is 0, 1, 5, 6, or 10 indicating the country for which the Nationally Determined Parameters (NDPs) are specified.

  0 = Other (NDPs are user specified)

  1 = CEN Default

  5 = Norway

  6 = Singapore

  10 = Portugal

EURO2004Direction

This is 1 or 2, indicating the ground motion direction.

1 = Horizontal

2 = Vertical (Does not apply when EURO2004Country = 6)

EURO2004SpectrumType

This is 1 or 2, indicating the spectrum type.

1 = Type 1

2 = Type 2 (Does not apply when EURO2004Country = 5 or 6)

EURO2004GroundType

This is 1, 2, 3, 4, 5 or 6, indicating the ground type.  This item only applies when the EURO2004Direction item is 1 (horizontal).

1 = A (Does not apply when EURO2004Country = 6)

2 = B (Does not apply when EURO2004Country = 6)

3 = C

4 = D

5 = E (Does not apply when EURO2004Country = 6)

6 = S1 (Only applies when EURO2004Country = 6)

EURO2004ag

The design ground acceleration in g, ag.

EURO2004S

The soil factor, S. This item only applies when the EURO2004Direction item is 1 (horizontal).

EURO2004AvgoverAg

The vertical ground acceleration divided by the horizontal ground acceleration, Avg/Ag. This item only applies when the EURO2004Direction item is 2 (vertical).

EURO2004Tb

The lower limit of period of the constant spectral acceleration branch, Tb.

EURO2004Tc

The upper limit of period of the constant spectral acceleration branch, Tc.

EURO2004Td

The period defining the start of the constant displacement range, Td.

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

      Dim SapObject as cOAPI

      Dim SapModel As cSapModel

      Dim ret As Long

      Dim EURO2004Country As Long

      Dim EURO2004Direction As Long

      Dim EURO2004SpectrumType As Long

      Dim EURO2004GroundType As Long

      Dim EURO2004ag As Double

      Dim EURO2004S As Double

      Dim EURO2004AvgoverAg As Double

      Dim EURO2004Tb As Double

      Dim EURO2004Tc As Double

      Dim EURO2004Td As Double

      Dim EURO2004Beta As Double

      Dim EURO2004q As Double

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



   'add Eurocode 8 2004 RS function

      ret = SapModel.Func.FuncRS.SetEurocode82004_1("RS-1", 1, 1, 1, 2, 0.4, 1.2, 0.9, 0.15, 0.5, 2, 0.2, 2, 0.04)



   'get Eurocode 8 2004 RS function

      ret = SapModel.Func.FuncRS.GetEurocode82004_1("RS-1", EURO2004Country, EURO2004Direction, EURO2004SpectrumType, EURO2004GroundType, EURO2004ag, EURO2004S, EURO2004AvgoverAg, EURO2004Tb, EURO2004Tc, EURO2004Td, EURO2004Beta, EURO2004q, DampRatio)



   'close Sap2000

      SapObject.ApplicationExit False

      Set SapModel = Nothing

      Set SapObject = Nothing

End Sub
```



## Release Notes

Initial release in version 14.1.0.

This function supersedes [GetEurocode82004](../../../Obsolete_Functions/GetEurocode82004_%28RS%29.md).

Added Portugal as a Country parameter in SAP2000 Version 15.0.0 and CSiBridge Version 15.1.0.

Added Singapore as a Country parameter in v20.1.0.

## See Also

[SetEurocode82004\_1](SetEurocode82004_1_%7BRS%7D.md)

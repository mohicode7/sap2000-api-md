# SetEurocode82004\_1

## Syntax

SapObject.SapModel.Func.FuncRS.SetEurocode82004\_1

## VB6 Procedure

Function SetEurocode82004\_1(ByVal Name As String, ByVal EURO2004Country As Long, ByVal EURO2004Direction As Long, ByVal EURO2004SpectrumType As Long, ByVal EURO2004GroundType As Long, ByVal EURO2004ag As Double, ByVal EURO2004S As Double, ByVal EURO2004AvgoverAg As Double, ByVal EURO2004Tb As Double, ByVal EURO2004Tc As Double, ByVal EURO2004Td As Double, ByVal EURO2004Beta As Double, ByVal EURO2004q As Double, ByVal DampRatio As Double) As Long

## Parameters

Name

The name of an existing or new function. If this is an existing function,n that function is modified; otherwise, a new function is added.

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

The soil factor, S. This item only applies when the EURO2004Direction item is 1 (horizontal).  If the EURO2004Country item is not 0, then the input value for this item is ignored.

EURO2004AvgoverAg

The vertical ground acceleration divided by the horizontal ground acceleration, Avg/Ag. This item only applies when the EURO2004Direction item is 2 (vertical).  If the EURO2004Country item is not 0, then the input value for this item is ignored.

EURO2004Tb

The lower limit of period of the constant spectral acceleration branch, Tb. If the EURO2004Country item is not 0, then the input value for this item is ignored.

EURO2004Tc

The upper limit of period of the constant spectral acceleration branch, Tc. If the EURO2004Country item is not 0, then the input value for this item is ignored.

EURO2004Td

The period defining the start of the constant displacement range, Td. If the EURO2004Country item is not 0, then the input value for this item is ignored.

EURO2004Beta

The lower bound factor, Beta. If the EURO2004Country item is not 0, then the input value for this item is ignored.

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



   'add Eurocode 8 2004 RS function

      ret = SapModel.Func.FuncRS.SetEurocode82004_1("RS-1", 1, 1, 1, 2, 0.4, 1.2, 0.9, 0.15, 0.5, 2, 0.2, 2, 0.04)



   'close Sap2000

      SapObject.ApplicationExit False

      Set SapModel = Nothing

      Set SapObject = Nothing

End Sub
```



## Release Notes

Initial release in version 14.1.0.

This function supersedes [SetEurocode82004](../../../Obsolete_Functions/SetEurocode82004_%7BRS%7D.md).

Added Portugal as a Country parameter in SAP2000 Version 15.0.0 and CSiBridge Version 15.1.0.

Added Singapore as a Country parameter in v20.1.0.

## See Also

[GetEurocode82004\_1](GetEurocode82004_1_%7BRS%7D.md)

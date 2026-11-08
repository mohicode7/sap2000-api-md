# SetNBCC2010

## Syntax

SapObject.SapModel.Func.FuncRS.SetNBCC2010

## VB6 Procedure

Function SetNBCC2010(ByVal Name As String, ByVal PGA As Double, ByVal S02 As Double, ByVal S05 As Double, ByVal S1 As Double, ByVal S2 As Double, ByVal SiteClass As Long, ByVal Fa As Double, ByVal Fv As Double, ByVal DampRatio As Double) As Long

## Parameters

Name

The name of an existing or new function. If this is an existing function, that function is modified; otherwise, a new function is added.

PGA

The peak ground acceleration.

S02

The spectral acceleration at a 0.2 second period.

S05

The spectral acceleration at a 0.5 second period.

S1

The spectral acceleration at a 1 second period.

S2

The spectral acceleration at a 2 second period.

SiteClass

This is 1, 2, 3, 4, 5 or 6, indicating the site class.

1 = A

2 = B

3 = C

4 = D

5 = E

6 = F

Fa

The site coefficient, Fa. This item is read when the site class is F only.

Fv

The site coefficient, Fv. This item is read when the site class is F only.

DampRatio

The damping ratio for the function, 0 <= DampRatio < 1.

## Remarks

This function defines a NBCC2010 response spectrum function.

The function returns zero if the function definition is successfully retrieved; otherwise it returns a nonzero value.

## VBA Example



```
Sub SetRSFuncNBCC2010()

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



   'add NBCC2010 RS function

      ret = SapModel.Func.FuncRS.SetNBCC2010("RS-1", 0.6, 1.1, 0.7, 0.35, 0.2, 6, 1.8, 2, 0.04)



   'close Sap2000

      SapObject.ApplicationExit False

      Set SapModel = Nothing

      Set SapObject = Nothing

End Sub
```



## Release Notes

Initial release in version 19.0.0.

## See Also

[GetNBCC2010](GetNBCC2010_%7BRS%7D.md)

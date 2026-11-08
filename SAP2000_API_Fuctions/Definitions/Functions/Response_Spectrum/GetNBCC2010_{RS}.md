# GetNBCC2010

## Syntax

SapObject.SapModel.Func.FuncRS.GetNBCC2010

## VB6 Procedure

Function GetNBCC2010(ByVal Name As String, ByRef PGA As Double, ByRef S02 As Double, ByRef S05 As Double, ByRef S1 As Double, ByRef S2 As Double, ByRef SiteClass As Long, ByRef Fa As Double, ByRef Fv As Double, ByRef DampRatio As Double) As Long

## Parameters

Name

The name of a NBCC2010 response spectrum function.

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

This function retrieves the definition of a NBCC2010 response spectrum function.

The function returns zero if the function definition is successfully retrieved; otherwise it returns a nonzero value.

## VBA Example



```
Sub GetRSFuncNBCC2010()

   'dimension variables

      Dim SapObject as cOAPI

      Dim SapModel As cSapModel

      Dim ret As Long

      Dim PGA As Double

      Dim S02 As Double
      Dim S05 As Double
      Dim S1 As Double
      Dim S2 As Double
      Dim SiteClass As Long

      Dim Fa As Double
      Dim Fv As Double

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



   'add NBCC2010 RS function

      ret = SapModel.Func.FuncRS.SetNBCC2010("RS-1", 0.6, 1.1, 0.7, 0.35, 0.2, 6, 1.8, 2, 0.04)



   'get NBCC2010 RS function

      ret = SapModel.Func.FuncRS.GetNBCC2010("RS-1", PGA, S02, S05, S1, S2, SiteClass, Fa, Fv, DampRatio)



   'close Sap2000

      SapObject.ApplicationExit False

      Set SapModel = Nothing

      Set SapObject = Nothing

End Sub
```



## Release Notes

Initial release in version 19.0.0.

## See Also

[SetNBCC2010](SetNBCC2010_%7BRS%7D.md)

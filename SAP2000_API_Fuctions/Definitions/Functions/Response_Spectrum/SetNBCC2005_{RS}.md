# SetNBCC2005

## Syntax

SapObject.SapModel.Func.FuncRS.SetNBCC2005

## VB6 Procedure

Function SetNBCC2005(ByVal Name As String, ByVal NBCC2005PGA As Double, ByVal NBCC2005S02 As Double, ByVal NBCC2005S05 As Double, ByVal NBCC2005S1 As Double, ByVal NBCC2005S2 As Double, ByVal NBCC2005SiteClass As Long, ByVal NBCC2005Fa As Double, ByVal NBCC2005Fv As Double, ByVal DampRatio As Double) As Long

## Parameters

Name

The name of an existing or new function. If this is an existing function, that function is modified; otherwise, a new function is added.

NBCC2005PGA

The peak ground acceleration.

NBCC2005S02

The spectral acceleration at a 0.2 second period.

NBCC2005S05

The spectral acceleration at a 0.52 second period.

NBCC2005S1

The spectral acceleration at a 1 second period.

NBCC2005S2

The spectral acceleration at a 2 second period.

NBCC2005SiteClass

This is 1, 2, 3, 4, 5 or 6, indicating the site class.

1 = A

2 = B

3 = C

4 = D

5 = E

6 = F

NBCC2005Fa

The site coefficient, Fa. This item is read when the site class is F only.

NBCC2005Fv

The site coefficient, Fv. This item is read when the site class is F only.

DampRatio

The damping ratio for the function, 0 <= DampRatio < 1.

## Remarks

This function defines a NBCC2005 response spectrum function.

The function returns zero if the function is successfully defined; otherwise it returns a nonzero value.

## VBA Example



```
Sub SetRSFuncNBCC2005()

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



   'add NBCC2005 RS function

      ret = SapModel.Func.FuncRS.SetNBCC2005("RS-1", 0.6, 1.1, 0.7, 0.35, 0.2, 6, 1.8, 2, 0.04)



   'close Sap2000

      SapObject.ApplicationExit False

      Set SapModel = Nothing

      Set SapObject = Nothing

End Sub
```



## Release Notes

Initial release in version 12.00.

## See Also

[GetNBCC2005](GetNBCC2005_%7BRS%7D.md)

# SetNBCC2015

## Syntax

SapObject.SapModel.Func.FuncRS.SetNBCC2015

## VB6 Procedure

Function SetNBCC2015(ByVal Name As String, ByVal PGA As Double, ByVal S02 As Double, ByVal S05 As Double, ByVal S1 As Double, ByVal S2 As Double, ByVal S5 As Double, ByVal S10 As Double, ByVal SiteClass As Long, ByVal F02 As Double, ByVal F05 As Double, ByVal F1 As Double, ByVal F2 As Double, ByVal F5 As Double, ByVal F10 As Double, ByVal DampRatio As Double) As Long

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

S5

The spectral acceleration at a 5 second period.

S10

The spectral acceleration at a 10 second period.

SiteClass

This is 1, 2, 3, 4, 5 or 6, indicating the site class.

1 = A

2 = B

3 = C

4 = D

5 = E

6 = F

F02

The site coefficient at a 0.2 second period. This item is read when the site class is F only.

F05

The site coefficient at a 0.5 second period. This item is read when the site class is F only.

F1

The site coefficient at a 1 second period. This item is read when the site class is F only.

F2

The site coefficient at a 2 second period. This item is read when the site class is F only.

F5

The site coefficient at a 5 second period. This item is read when the site class is F only.

F10

The site coefficient at a 10 second period. This item is read when the site class is F only.

DampRatio

The damping ratio for the function, 0 <= DampRatio < 1.

## Remarks

This function retrieves the definition of a NBCC2015 response spectrum function.

The function returns zero if the function definition is successfully retrieved; otherwise it returns a nonzero value.

## VBA Example



```
Sub SetRSFuncNBCC2015()

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



   'add NBCC2015 RS function

      ret = SapModel.Func.FuncRS.SetNBCC2015("RS-1", 0.6, 1.1, 0.7, 0.35, 0.2, 0.035, 0.01, 6, 1.1, 1.2, 1.3, 1.4, 1.5, 1.6, 0.04)



    'close Sap2000

      SapObject.ApplicationExit False

      Set SapModel = Nothing

      Set SapObject = Nothing

End Sub
```



## Release Notes

Initial release in version 19.0.0.

## See Also

[GetNBCC2015](GetNBCC2015_%7BRS%7D.md)

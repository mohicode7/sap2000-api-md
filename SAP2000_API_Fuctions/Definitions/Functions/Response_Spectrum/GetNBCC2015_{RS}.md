# GetNBCC2015

## Syntax

SapObject.SapModel.Func.FuncRS.GetNBCC2015

## VB6 Procedure

Function GetNBCC2015(ByVal Name As String, ByRef PGA As Double, ByRef S02 As Double, ByRef S05 As Double, ByRef S1 As Double, ByRef S2 As Double, ByRef S5 As Double, ByRef S10 As Double, ByRef SiteClass As Long, ByRef F02 As Double, ByRef F05 As Double, ByRef F1 As Double, ByRef F2 As Double, ByRef F5 As Double, ByRef F10 As Double, ByRef DampRatio As Double) As Long

## Parameters

Name

The name of a NBCC2015 response spectrum function.

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
Sub GetRSFuncNBCC2015()

   'dimension variables

      Dim SapObject as cOAPI

      Dim SapModel As cSapModel

      Dim ret As Long

      Dim PGA As Double

      Dim S02 As Double
      Dim S05 As Double
      Dim S1 As Double
      Dim S2 As Double
      Dim S5 As Double
      Dim S10 As Double
      Dim SiteClass As Long

      Dim F02 As Double
      Dim F05 As Double
      Dim F1 As Double
      Dim F2 As Double
      Dim F5 As Double
      Dim F10 As Double

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



   'add NBCC2015 RS function

      ret = SapModel.Func.FuncRS.SetNBCC2015("RS-1", 0.6, 1.1, 0.7, 0.35, 0.2, 0.035, 0.01, 6, 1.1, 1.2, 1.3, 1.4, 1.5, 1.6, 0.04)



   'get NBCC2015 RS function

      ret = SapModel.Func.FuncRS.GetNBCC2015("RS-1", PGA, S02, S05, S1, S2, S5, S10, SiteClass, F02, F05, F1, F2, F5, F10, DampRatio)



   'close Sap2000

      SapObject.ApplicationExit False

      Set SapModel = Nothing

      Set SapObject = Nothing

End Sub
```



## Release Notes

Initial release in version 19.0.0.

## See Also

[SetNBCC2015](SetNBCC2015_%7BRS%7D.md)

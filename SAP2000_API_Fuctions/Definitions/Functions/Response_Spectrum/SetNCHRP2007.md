# SetNCHRP2007

## Syntax

SapObject.SapModel.Func.FuncRS.SetNCHRP2007

## VB6 Procedure

Function SetNCHRP2007(ByVal Name As String, ByVal NCHRP2007Option As Long, ByVal NCHRP2007Latitude As Double, ByVal NCHRP2007Longitude As Double, ByVal NCHRP2007ZipCode As String, ByVal NCHRP2007SS As Double, ByVal NCHRP2007S1 As Double, ByVal NCHRP2007SiteClass As Long, ByVal NCHRP2007Fa As Double, ByRef NCHRP2007Fv As Double, ByVal DampRatio As Double) As Long

## Parameters

Name

The name of an existing or new function. If this is an existing function, that function is modified; otherwise, a new function is added.

NCHRP2007Option

This is 0, 1 or 2, indicating the seismic coefficient option.

0 = Ss and S1 from USGS by latitiude and longitude

1 = Ss and S1 from USGS by zip code

2 = Ss and S1 are user defined

NCHRP2007Latitude, NCHRP2007Longitude

The latitude and longitude for which the seismic coefficients are obtained. These items are used only when NCHRP2007Option = 0.

NCHRP2007ZipCode

The zip code for which the seismic coefficients are obtained. This item is used only when NCHRP2007Option = 1.

NCHRP2007SS, NCHRP2007S1

The seismic coefficients Ss and S1. This item is used only when NCHRP2007Option = 2.

NCHRP2007SiteClass

This is 1, 2, 3, 4, 5 or 6, indicating the site class.

1 = A

2 = B

3 = C

4 = D

5 = E

6 = F

NCHRP2007Fa, NCHRP2007Fv

The site coefficients Fa and Fv. These items are used only when NCHRP2007SiteClass= 6.

DampRatio

The damping ratio for the function, 0 <= DampRatio < 1.

## Remarks

This function defines a NCHRP 20-07 response spectrum function.

The function returns zero if the function is successfully defined; otherwise it returns a nonzero value.

## VBA Example



```
Sub SetRSFuncNCHRP2007()

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



   'add NCHRP2007 RS function

      ret = SapModel.Func.FuncRS.SetNCHRP2007("RS-1", 1, 0, 0, "94704", 0, 0, 4, 0, 0, 0.04)



   'close Sap2000

      SapObject.ApplicationExit False

      Set SapModel = Nothing

      Set SapObject = Nothing

End Sub
```



## Release Notes

Initial release in version 11.02.

## See Also

[GetNCHRP2007](GetNCHRP2007.md)

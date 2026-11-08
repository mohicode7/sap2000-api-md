# SetAASHTO2007

## Syntax

SapObject.SapModel.Func.FuncRS.SetAASHTO2007

## VB6 Procedure

Function SetAASHTO2007(ByVal Name As String, ByVal AASHTO2007Option As Long, ByVal AASHTO2007Latitude As Double, ByVal AASHTO2007Longitude As Double, ByVal AASHTO2007ZipCode As String, ByVal AASHTO2007SS As Double, ByVal AASHTO2007S1 As Double, ByVal AASHTO2007PGA As Double, ByVal AASHTO2007SiteClass As Long, ByVal AASHTO2007Fa As Double, ByRef AASHTO2007Fv As Double, ByVal AASHTO2007Fpga As Double, ByVal DampRatio As Double) As Long

## Parameters

Name

The name of an existing or new function. If this is an existing function, that function is modified; otherwise, a new function is added.

AASHTO2007Option

This is 0, 1 or 2, indicating the seismic coefficient option.

0 = Ss and S1 from USGS by latitude and longitude

1 = Ss and S1 from USGS by zip code

2 = Ss and S1 are user defined

AASHTO2007Latitude, AASHTO2007Longitude

The latitude and longitude for which the seismic coefficients are obtained. These items are used only when AASHTO2007Option = 0.

AASHTO2007ZipCode

The zip code for which the seismic coefficients are obtained. This item is used only when AASHTO2007Option = 1.

AASHTO2007SS, AASHTO2007S1, AASHTO2007PGA

The seismic coefficients Ss, S1 and PGA. These items are used only when AASHTO2007Option = 2.

AASHTO2007SiteClass

This is 1, 2, 3, 4, 5 or 6, indicating the site class.

1 = A

2 = B

3 = C

4 = D

5 = E

6 = F

AASHTO2007Fa, AASHTO2007Fv, AASHTO2007Fpga

The site coefficients Fa, Fv and Fpga. These items are used only when AASHTO2007SiteClass= 6.

DampRatio

The damping ratio for the function, 0 <= DampRatio < 1.

## Remarks

This function defines a AASHTO 2007 response spectrum function.

The function returns zero if the function is successfully defined; otherwise it returns a nonzero value.

## VBA Example



```
Sub SetRSFuncAASHTO2007()

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



   'add AASHTO2007 RS function

      ret = SapModel.Func.FuncRS.SetAASHTO2007("RS-1", 1, 0, 0, "94704", 0, 0, 0, 4, 0, 0, 0, 0.04)



   'close Sap2000

      SapObject.ApplicationExit False

      Set SapModel = Nothing

      Set SapObject = Nothing

End Sub
```



## Release Notes

Initial release in version 14.00.

## See Also

[GetAASHTO2007](GetAASHTO2007.md)

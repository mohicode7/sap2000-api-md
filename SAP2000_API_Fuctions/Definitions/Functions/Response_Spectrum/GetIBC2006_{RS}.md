# GetIBC2006

## Syntax

SapObject.SapModel.Func.FuncRS.GetIBC2006

## VB6 Procedure

Function GetIBC2006(ByVal Name As String, ByRef IBC2006Option As Long, ByRef IBC2006Latitude As Double, ByRef IBC2006Longitude As Double, ByRef IBC2006ZipCode As String, ByRef IBC2006SS As Double, ByRef IBC2006S1 As Double, ByRef IBC2006TL As Double, ByRef IBC2006SiteClass As Long, ByRef IBC2006Fa As Double, ByRef IBC2006Fv As Double, ByRef DampRatio As Double) As Long

## Parameters

Name

The name of a IBC2006 response spectrum function.

IBC2006Option

This is 0, 1 or 2, indicating the seismic coefficient option.

0 = Ss and S1 from USGS by latitiude and longitude

1 = Ss and S1 from USGS by zip code

2 = Ss and S1 are user defined

IBC2006Latitude, IBC2006Longitude

The latitude and longitude for which the seismic coefficients are obtained. These items are used only when IBC2006Option = 0.

IBC2006ZipCode

The zip code for which the seismic coefficients are obtained. This item is used only when IBC2006Option = 1.

IBC2006SS, IBC2006S1

The seismic coefficients Ss and S1. This item is used only when IBC2006Option = 2.

IBC2006TL

The long-period transition period. [s]

IBC2006SiteClass

This is 1, 2, 3, 4, 5 or 6, indicating the site class.

1 = A

2 = B

3 = C

4 = D

5 = E

6 = F

IBC2006Fa, IBC2006Fv

The site coefficients Fa and Fv. These items are used only when IBC2006SiteClass= 6.

DampRatio

The damping ratio for the function, 0 <= DampRatio < 1.

## Remarks

This function retrieves the definition of a IBC2006 response spectrum function.

The function returns zero if the function definition is successfully retrieved; otherwise it returns a nonzero value.

## VBA Example



```
Sub GetRSFuncIBC2006()

   'dimension variables

      Dim SapObject as cOAPI

      Dim SapModel As cSapModel

      Dim ret As Long

      Dim IBC2006Option As Long

      Dim IBC2006Latitude As Double

      Dim IBC2006Longitude As Double

      Dim IBC2006ZipCode As String

      Dim IBC2006SS As Double

      Dim IBC2006S1 As Double

      Dim IBC2006TL As Double

      Dim IBC2006SiteClass As Long

      Dim IBC2006Fa As Double

      Dim IBC2006Fv As Double

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



   'add IBC2006 RS function

      ret = SapModel.Func.FuncRS.SetIBC2006("RS-1", 1, 0, 0, "94704", 0, 0, 7.5, 4, 0, 0, 0.04)



   'get IBC2006 RS function

      ret = SapModel.Func.FuncRS.GetIBC2006("RS-1", IBC2006Option, IBC2006Latitude, IBC2006Longitude, IBC2006ZipCode, IBC2006SS, IBC2006S1, IBC2006TL, IBC2006SiteClass, IBC2006Fa, IBC2006Fv, DampRatio)



   'close Sap2000

      SapObject.ApplicationExit False

      Set SapModel = Nothing

      Set SapObject = Nothing

End Sub
```



## Release Notes

Initial release in version 11.02.

## See Also

[SetIBC2006](SetIBC2006_%7BRS%7D.md)

# SetLoads

## Syntax

SapObject.SapModel.LoadCases.ResponseSpectrum.SetLoads

## VB6 Procedure

Function SetLoads(ByVal Name As String, ByVal NumberLoads As Long, ByRef LoadName() As String, ByRef Func() As String, ByRef SF() As Double, ByRef CSys() As String, ByRef Ang() As Double) As Long

## Parameters

Name

The name of an existing response spectrum load case.

NumberLoads

The number of loads assigned to the specified analysis case.

LoadName

This is an array that includes U1, U2, U3, R1, R2 or R3, indicating the direction of the load.

Func

This is an array that includes the name of the response spectrum function associated with each load.

SF

This is an array that includes the scale factor of each load assigned to the load case. [L/s2] for U1 U2 and U3; otherwise unitless

CSys

This is an array that includes the name of the coordinate system associated with each load. If this item is a blank string, the Global coordinate system is assumed.

Ang

This is an array that includes the angle between the acceleration local 1 axis and the +X-axis of the coordinate system specified by the CSys item. The rotation is about the Z-axis of the specified coordinate system. [deg]

## Remarks

This function sets the load data for the specified analysis case.

The function returns zero if the data is successfully set; otherwise it returns a nonzero value.

## VBA Example



```
Sub SetCaseResponseSpectrumLoads()

   'dimension variables

      Dim SapObject as cOAPI

      Dim SapModel As cSapModel

      Dim ret As Long

      Dim MyLoadName() As String

      Dim MyFunc() As String

      Dim MySF() As Double

      Dim MyCSys() As String

      Dim MyAng() As Double



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



   'add UBC97 RS function

      ret = SapModel.Func.FuncRS.SetUBC97("RS-1", 0.36, 0.54, 0.04)



   'add response spectrum load case

      ret = SapModel.LoadCases.ResponseSpectrum.SetCase("LCASE1")



   'set load data

      ReDim MyLoadName(1)

      ReDim MyFunc(1)

      ReDim MySF(1)

      ReDim MyCSys(1)

      ReDim MyAng(1)

      MyLoadName(0) = "U1"

      MyFunc(0) = "RS-1"

      MySF(0) = 386

      MyCSys(0) = "Global"

      MyAng(0) = 10

      MyLoadName(1) = "U2"

      MyFunc(1) = "RS-1"

      MySF(1) = 386

      ret = SapModel.LoadCases.ResponseSpectrum.SetLoads("LCASE1", 2, MyLoadName, MyFunc, MySF, MyCSys, MyAng)



   'close Sap2000

      SapObject.ApplicationExit False

      Set SapModel = Nothing

      Set SapObject = Nothing

End Sub
```



## Release Notes

Initial release in version 11.02.

Changed nomenclature from Load Cases, Analysis Cases and Response Combinations to Load Patterns, Load Cases and Load Combinations, respectively, in version 12.00.

## See Also

[GetLoads](GetLoads_%7BResponse_Spectrum%7D.md)

# GetLoads

## Syntax

SapObject.SapModel.LoadCases.SteadyState.GetLoads

## VB6 Procedure

Function GetLoads(ByVal Name As String, ByRef NumberLoads As Long, ByRef LoadType() As String, ByRef LoadName() As String, ByRef Func() As String, ByRef SF() As Double, ByRef PhaseAngle() As Double, ByRef CSys() As String, ByRef Ang() As Double) As Long

## Parameters

Name

The name of an existing steady state load case.

NumberLoads

The number of loads assigned to the specified analysis case.

LoadType

This is an array that includes either Load or Accel, indicating the type of each load assigned to the load case.

LoadName

This is an array that includes the name of each load assigned to the load case.

If the LoadType item is Load, this item is the name of a defined load pattern.

If the LoadType item is Accel, this item is U1, U2, U3, R1, R2 or R3, indicating the direction of the load.

Func

This is an array that includes the name of the steady state function associated with each load.

SF

This is an array that includes the scale factor of each load assigned to the load case. [L/s2] for U1 U2 and U3; otherwise unitless

PhaseAngle

This is an array that includes the phase angle. [deg]

CSys

This is an array that includes the name of the coordinate system associated with each load. If this item is a blank string, the Global coordinate system is assumed.

This item applies only when the LoadType item is Accel.

Ang

This is an array that includes the angle between the acceleration local 1 axis and the +X-axis of the coordinate system specified by the CSys item. The rotation is about the Z-axis of the specified coordinate system. [deg]

This item applies only when the LoadType item is Accel.

## Remarks

This function retrieves the load data for the specified load case.

The function returns zero if the data is successfully retrieved; otherwise it returns a nonzero value.

## VBA Example



```
Sub GetCaseSteadyStateLoads()

   'dimension variables

      Dim SapObject as cOAPI

      Dim SapModel As cSapModel

      Dim ret As Long

      Dim MyLoadType() As String

      Dim MyLoadName() As String

      Dim MyFunc() As String

      Dim MySF() As Double

      Dim MyPhaseAngle() As Double

      Dim MyCSys() As String

      Dim MyAng() As Double

      Dim NumberLoads As Long

      Dim LoadType() As String

      Dim LoadName() As String

      Dim Func() As String

      Dim SF() As Double

      Dim PhaseAngle() As Double

      Dim CSys() As String

      Dim Ang() As Double



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



   'add steady state load case

      ret = SapModel.LoadCases.SteadyState.SetCase("LCASE1")



   'set load data

      ReDim MyLoadType(1)

      ReDim MyLoadName(1)

      ReDim MyFunc(1)

      ReDim MySF(1)

      ReDim MyPhaseAngle(1)

      ReDim MyCSys(1)

      ReDim MyAng(1)

      MyLoadType(0) = "Load"

      MyLoadName(0) = "DEAD"

      MyFunc(0) = "UNIFSS"

      MySF(0) = 1.5

      MyPhaseAngle(0) = 90

      MyLoadType(1) = "Accel"

      MyLoadName(1) = "U3"

      MyFunc(1) = "UNIFSS"

      MySF(1) = 1

      MyPhaseAngle(1) = 90

      MyCSys(1) = "Global"

      MyAng(1) = 10

      ret = SapModel.LoadCases.SteadyState.SetLoads("LCASE1", 2, MyLoadType, MyLoadName, MyFunc, MySF, MyPhaseAngle, MyCSys, MyAng)



   'get load data

      ret = SapModel.LoadCases.SteadyState.GetLoads("LCASE1", NumberLoads, LoadType, LoadName, Func, SF, PhaseAngle, CSys, Ang)



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

[SetLoads](SetLoads_%7BSteady_State%7D.md)

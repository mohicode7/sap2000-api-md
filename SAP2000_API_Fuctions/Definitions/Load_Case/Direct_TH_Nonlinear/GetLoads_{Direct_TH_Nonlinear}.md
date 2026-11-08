# GetLoads

## Syntax

SapObject.SapModel.LoadCases.DirHistNonlinear.GetLoads

## VB6 Procedure

Function GetLoads(ByVal Name As String, ByRef NumberLoads As Long, ByRef LoadType() As String, ByRef LoadName() As String, ByRef Func() As String, ByRef SF() As Double, ByRef TF() As Double, ByRef AT() As Double, ByRef CSys() As String, ByRef Ang() As Double) As Long

## Parameters

Name

The name of an existing nonlinear direct integration time history load case.

NumberLoads

The number of loads assigned to the specified analysis case.

LoadType

This is an array that includes Load or Accel, indicating the type of each load assigned to the load case.

LoadName

This is an array that includes the name of each load assigned to the load case.

If the LoadType item is Load, this item is the name of a defined load pattern.

If the LoadType item is Accel, this item is U1, U2, U3, R1, R2 or R3, indicating the direction of the load.

Func

This is an array that includes the name of the time history function associated with each load.

SF

This is an array that includes the scale factor of each load assigned to the load case. [L/s2] for U1 U2 and U3; otherwise unitless

TF

This is an array that includes the time scale factor of each load assigned to the load case.

AT

This is an array that includes the arrival time of each load assigned to the load case.

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
Sub GetCaseDirHistNonlinearLoads()

   'dimension variables

      Dim SapObject as cOAPI

      Dim SapModel As cSapModel

      Dim ret As Long

      Dim MyLoadType() As String

      Dim MyLoadName() As String

      Dim MyFunc() As String

      Dim MySF() As Double

      Dim MyTF() As Double

      Dim MyAT() As Double

      Dim MyCSys() As String

      Dim MyAng() As Double

      Dim NumberLoads As Long

      Dim LoadType() As String

      Dim LoadName() As String

      Dim Func() As String

      Dim SF() As Double

      Dim TF() As Double

      Dim AT() As Double

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



   'add sine TH function

      ret = SapModel.Func.FuncTH.SetSine("TH-1", 1, 16, 4, 1.25)



   'add nonlinear direct history load case

      ret = SapModel.LoadCases.DirHistNonlinear.SetCase("LCASE1")



   'set load data

      ReDim MyLoadType(1)

      ReDim MyLoadName(1)

      ReDim MyFunc(1)

      ReDim MySF(1)

      ReDim MyTF(1)

      ReDim MyAT(1)

      ReDim MyCSys(1)

      ReDim MyAng(1)

      MyLoadType(0) = "Load"

      MyLoadName(0) = "DEAD"

      MyFunc(0) = "RAMPTH"

      MySF(0) = 1

      MyTF(0) = 1

      MyAT(0) = 0

      MyCSys(0) = "Global"

      MyAng(0) = 0

      MyLoadType(1) = "Accel"

      MyLoadName(1) = "U2"

      MyFunc(1) = "TH-1"

      MySF(1) = 2

      MyTF(1) = 1.5

      MyAT(1) = 10

      MyCSys(1) = "Global"

      MyAng(1) = 10

      ret = SapModel.LoadCases.DirHistNonlinear.SetLoads("LCASE1", 2, MyLoadType, MyLoadName, MyFunc, MySF, MyTF, MyAT, MyCSys, MyAng)



   'get load data

      ret = SapModel.LoadCases.DirHistNonlinear.GetLoads("LCASE1", NumberLoads, LoadType, LoadName, Func, SF, TF, AT, CSys, Ang)



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

[SetLoads](SetLoads_%7BDirect_TH_Nonlinear%7D.md)

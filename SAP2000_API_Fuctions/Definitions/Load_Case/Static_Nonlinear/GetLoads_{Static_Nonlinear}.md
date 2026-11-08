# GetLoads

## Syntax

SapObject.SapModel.LoadCases.StaticNonlinear.GetLoads

## VB6 Procedure

Function GetLoads(ByVal Name As String, ByRef NumberLoads As Long, ByRef LoadType() As String, ByRef LoadName() As String, ByRef SF() As Double) As Long

## Parameters

Name

The name of an existing static nonlinear load case.

NumberLoads

The number of loads assigned to the specified analysis case.

LoadType

This is an array that includes either Load or Accel, indicating the type of each load assigned to the load case.

LoadName

This is an array that includes the name of each load assigned to the load case.

If the LoadType item is Load, this item is the name of a defined load pattern.

If the LoadType item is Accel, this item is UX, UY, UZ, RX, RY or RZ, indicating the direction of the load.

SF

This is an array that includes the scale factor of each load assigned to the load case. [L/s2] for Accel UX UY and UZ; otherwise unitless

## Remarks

This function retrieves the load data for the specified load case.

The function returns zero if the data is successfully retrieved; otherwise it returns a nonzero value.

## VBA Example



```
Sub GetCaseStaticNonlinearLoads()

   'dimension variables

      Dim SapObject as cOAPI

      Dim SapModel As cSapModel

      Dim ret As Long

      Dim MyLoadType() As String

      Dim MyLoadName() As String

      Dim MySF() As Double

      Dim NumberLoads As Long

      Dim LoadType() As String

      Dim LoadName() As String

      Dim SF() As Double



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



   'add static nonlinear load case

      ret = SapModel.LoadCases.StaticNonlinear.SetCase("LCASE1")



   'set load data

      ReDim MyLoadType(1)

      ReDim MyLoadName(1)

      ReDim MySF(1)

      MyLoadType(0) = "Load"

      MyLoadName(0) = "DEAD"

      MySF(0) = 0.7

      MyLoadType(1) = "Accel"

      MyLoadName(1) = "UZ"

      MySF(1) = 1.2

      ret = SapModel.LoadCases.StaticNonlinear.SetLoads("LCASE1", 2, MyLoadType, MyLoadName, MySF)



   'get load data

      ret = SapModel.LoadCases.StaticNonlinear.GetLoads("LCASE1", NumberLoads, LoadType, LoadName, SF)



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

[SetLoads](SetLoads_%7BStatic_Nonlinear%7D.md)

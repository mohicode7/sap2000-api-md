# GetLoads\_1

## Syntax

SapObject.SapModel.LoadCases.StaticLinearMultistep.GetLoads\_1

## VB6 Procedure

Function GetLoads\_1(ByVal Name As String, ByRef NumberLoads As Long, ByRef LoadType() As String, ByRef LoadName() As String, ByRef SF() As Double, ByRef StepRange() as Integer, ByRef FirstLoadStep() As Integer, ByRef LastLoadStep() As Integer, ByRef StartCaseStep() As Integer, ByRef ExtrapolateOption() As Integer) As Long

## Parameters

Name

The name of an existing static linear multistep analysis case.

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

StepRange

This is an array that identifies the step range type to consider for each load assigned to the load case. The allowed values are:

0 = All

1 = User

FirstLoadStep

This is an array that specifies the first load step to consider for each load assigned to the load case. This value is only applicable when StepRange = User.

LastLoadStep

This is an array that specifies the last load step to consider for each load assigned to the load case. This value is only applicable when StepRange = User.

StartCaseStep

This is an array that specifies the load case step at which to start applying each load assigned to the load case.

ExtrapolateOption

This is an array that identifies the extrapolation option for each load assigned to the load case. The allowed values are:

0 = None

1 = Last Step

2 = Repeat Range

## Remarks

This function retrieves the load data for the specified load case.

The function returns zero if the data is successfully retrieved; otherwise, it returns a nonzero value.

## VBA Example



```
Sub GetCaseStaticLinearMultistepLoads()

   'dimension variables

      Dim SapObject as cOAPI

      Dim SapModel As cSapModel

      Dim ret As Long

      Dim MyLoadType() As String

      Dim MyLoadName() As String

      Dim MySF() As Double

      Dim MyStepRange() As Long
      Dim MyFirstLoadStep() As Long
      Dim MyLastLoadStep() As Long
      Dim MyStartCaseStep() As Long
      Dim MyExtrapolateOption() As Long
      Dim NumberLoads As Long

      Dim LoadType() As String

      Dim LoadName() As String

      Dim SF() As Double
      Dim StepRange() As Long
      Dim FirstLoadStep() As Long
      Dim LastLoadStep() As Long
      Dim StartCaseStep() As Long
      Dim ExtrapolateOption() As Long



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



   'add static linear multistep load case

      ret = SapModel.LoadCases.StaticLinearMultistep.SetCase_1("LCASE1")



   'set load data

      ReDim MyLoadType(1)

      ReDim MyLoadName(1)

      ReDim MySF(1)
      ReDim MyStepRange(1)
      ReDim MyFirstLoadStep(1)
      ReDim MyLastLoadStep(1)
      ReDim MyStartCaseStep(1)
      ReDim MyExtrapolateOption(1)

      MyLoadType(0) = "Load"

      MyLoadName(0) = "DEAD"

      MySF(0) = 0.7
      MyStepRange(0) = 0
      MyFirstLoadStep(0) = 1
      MyLastLoadStep(0) = 1

      MyStartCaseStep(0) = 1
      MyExtrapolateOption(0) = 1
      MyLoadType(1) = "Accel0"

      MyLoadName(1) = "UZ"

      MySF(1) = 1.2

      ret = SapModel.LoadCases.StaticLinearMultistep.SetLoads_1("LCASE1", 2, MyLoadType, MyLoadName, MySF, MyStepRange, MyFirstLoadStep, MyLastLoadStep, MyStartCaseStep, MyExtrapolateOption)



   'get load data

      ret = SapModel.LoadCases.StaticLinearMultistep.GetLoads_1("LCASE1", NumberLoads, LoadType, LoadName, SF, StepRange, FirstLoadStep, LastLoadStep, StartCaseStep, ExtrapolateOption)



   'close Sap2000

      SapObject.ApplicationExit False

      Set SapModel = Nothing

      Set SapObject = Nothing

End Sub
```



## Release Notes

Initial release in v21.0.0.

This function supersedes [GetLoads](../../../Obsolete_Functions/GetLoads_%7BStatic_Linear_Multistep%7D.md).

## See Also

[SetLoads\_1](SetLoads_%7BStatic_Linear_Multistep%7D_1.md)

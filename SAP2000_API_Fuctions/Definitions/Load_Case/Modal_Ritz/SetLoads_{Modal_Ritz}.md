# SetLoads

## Syntax

SapObject.SapModel.LoadCases.ModalRitz.SetLoads

## VB6 Procedure

Function SetLoads(ByVal Name As String, ByVal NumberLoads As Long, ByRef LoadType() As String, ByRef LoadName() As String, ByRef RitzMaxCyc() As Long, ByRef TargetPar() As Double) As Long

## Parameters

Name

The name of an existing modal ritz load case.

NumberLoads

The number of loads assigned to the specified analysis case.

LoadType

This is an array that includes Load, Accel or Link, indicating the type of each load assigned to the load case.

LoadName

This is an array that includes the name of each load assigned to the load case.

If the LoadType item is Load, this item is the name of a defined load pattern.

If the LoadType item is Accel, this item is UX, UY, UZ, RX, RY or RZ, indicating the direction of the load.

If the LoadType item is Link, this item is not used.

RitzMaxCyc

This is an array that includes the maximum number of generation cycles to be performed for the specified ritz starting vector. A value of 0 means there is no limit on the number of cycles.

TargetPar

This is an array that includes the target dynamic participation ratio.

## Remarks

This function sets the load data for the specified analysis case.

The function returns zero if the data is successfully set; otherwise it returns a nonzero value.

## VBA Example



```
Sub SetCaseModalRitzLoads()

   'dimension variables

      Dim SapObject as cOAPI

      Dim SapModel As cSapModel

      Dim ret As Long

      Dim MyLoadType() As String

      Dim MyLoadName() As String

      Dim MyRitzMaxCyc() As Long

      Dim MyTargetPar() As Double



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



   'add modal ritz load case

      ret = SapModel.LoadCases.ModalRitz.SetCase("LCASE1")



   'set load data

      ReDim MyLoadType(2)

      ReDim MyLoadName(2)

      ReDim MyRitzMaxCyc(2)

      ReDim MyTargetPar(2)

      MyLoadType(0) = "Load"

      MyLoadName(0) = "DEAD"

      MyRitzMaxCyc(0) = 0

      MyTargetPar(0) = 99

      MyLoadType(1) = "Accel"

      MyLoadName(1) = "UZ"

      MyRitzMaxCyc(1) = 0

      MyTargetPar(1) = 99

      MyLoadType(2) = "Link"

      MyRitzMaxCyc(2) = 0

      MyTargetPar(2) = 99

      ret = SapModel.LoadCases.ModalRitz.SetLoads("LCASE1", 3, MyLoadType, MyLoadName, MyRitzMaxCyc, MyTargetPar)



   'close Sap2000

      SapObject.ApplicationExit False

      Set SapModel = Nothing

      Set SapObject = Nothing

End Sub
```



## Release Notes

Initial release in version 11.03.

Changed nomenclature from Load Cases, Analysis Cases and Response Combinations to Load Patterns, Load Cases and Load Combinations, respectively, in version 12.00.

## See Also

[GetLoads](GetLoads_%7BModal_Ritz%7D.md)

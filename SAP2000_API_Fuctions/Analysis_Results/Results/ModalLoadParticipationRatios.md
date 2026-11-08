# ModalLoadParticipationRatios

## Syntax

SapObject.SapModel.Results.ModalLoadParticipationRatios

## VB6 Procedure

Function ModalLoadParticipationRatios(ByRef NumberResults As Long, ByRef LoadCase() As String, ByRef ItemType() As String, ByRef Item() As String, ByRef Stat() As Double, ByRef Dyn() As Double) As Long

## Parameters

NumberResults

The total number of results returned by the program.

LoadCase

This is an array that includes the name of the modal load case associated with each result.

ItemType

This is an array that includes Load Pattern, Acceleration, Link or Panel Zone. It specifies the type of item for which the modal load participation is reported.

Item

This is an array whose values depend on the ItemType. If the ItemType is Load Pattern, this is the name of the load pattern.

If the ItemType is Acceleration, this is UX, UY, UZ, RX, RY, or RZ, indicating the acceleration direction.

If the ItemType is Link, this is the name of the link followed by U1, U2, U3, R1, R2, or R3 (in parenthesis), indicating the link degree of freedom for which the output is reported.

If the ItemType is Panel Zone, this is the name of the joint to which the panel zone is assigned, followed by U1, U2, U3, R1, R2, or R3 (in parenthesis), indicating the degree of freedom for which the output is reported.

Stat

This is an array that includes the percent static load participation ratio.

Dyn

This is an array that includes the percent dynamic load participation ratio.

## Remarks

This function reports the modal load participation ratios for each selected modal analysis case.

The function returns zero if the data is successfully recovered; otherwise it returns a nonzero value.

See [Analysis Results Remarks](Analysis_Results_Remarks.md) for more information.

## VBA Example



```
Sub GetModalLoadParticipationRatios()

   'dimension variables

      Dim SapObject as cOAPI

      Dim SapModel As cSapModel

      Dim ret As Long

      Dim NumberResults As Long

      Dim LoadCase() As String

      Dim ItemType() As String

      Dim Item() As String

      Dim Stat() As Double

      Dim Dyn() As Double



   'create Sap2000 object

      Set SapObject = CreateObject("CSI.SAP2000.API.SapObject")



   'start Sap2000 application

      SapObject.ApplicationStart



   'create SapModel object

      Set SapModel = SapObject.SapModel



   'initialize model

      ret = SapModel.InitializeNewModel



   'create model from template

      ret = SapModel.File.New2DFrame(PortalFrame, 3, 124, 3, 200)



   'run analysis

      ret = SapModel.File.Save("C:\SapAPI\x.sdb")

      ret = SapModel.Analyze.RunAnalysis



   'clear all case and combo output selections

      ret = SapModel.Results.Setup.DeselectAllCasesAndCombosForOutput



   'set case and combo output selections

      ret = SapModel.Results.Setup.SetCaseSelectedForOutput("MODAL")



   'get modal load participation ratios

      ret = SapModel.Results.ModalLoadParticipationRatios(NumberResults, LoadCase, ItemType, Item, Stat, Dyn)



   'close Sap2000

      SapObject.ApplicationExit False

      Set SapModel = Nothing

      Set SapObject = Nothing

   End Sub
```



## Release Notes

Initial release in version 12.00.

## See Also

[ModeShape](ModeShape.md)

[ModalPeriod](ModalPeriod.md)

[ModalParticipatingMassRatios](ModalParticipatingMassRatios.md)

[ModalParticipationFactors](ModalParticipationFactors.md)

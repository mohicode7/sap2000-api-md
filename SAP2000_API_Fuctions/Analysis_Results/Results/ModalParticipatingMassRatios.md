# ModalParticipatingMassRatios

## Syntax

SapObject.SapModel.Results.ModalParticipatingMassRatios

## VB6 Procedure

Function ModalParticipatingMassRatios(ByRef NumberResults As Long, ByRef LoadCase() As String, ByRef StepType() As String, ByRef StepNum() As Double, ByRef Period() As Double, ByRef Ux() As Double, ByRef Uy() As Double, ByRef Uz() As Double, ByRef SumUx() As Double, ByRef SumUy() As Double, ByRef SumUz() As Double, ByRef Rx() As Double, ByRef Ry() As Double, ByRef Rz() As Double, ByRef SumRx() As Double, ByRef SumRy() As Double, ByRef SumRz() As Double) As Long

## Parameters

NumberResults

The total number of results returned by the program.

LoadCase

This is an array that includes the name of the modal load case associated with each result.

StepType

This is an array that includes the step type, if any, for each result. For modal results, this will always be Mode.See [Analysis Results Step Type](Analysis_Results_Step_Type.md).

StepNum

This is an array that includes the step number for each result. For modal results, this is always the mode number. See [Analysis Results Step Number](Analysis_Results_Step_Number.md).

Period

This is an array that includes the period for each result. [s]

Ux

This is an array that includes the modal participating mass ratio for the structure Ux degree of freedom. The ratio applies to the specified mode.

Uy

This is an array that includes the modal participating mass ratio for the structure Uy degree of freedom. The ratio applies to the specified mode.

Uz

This is an array that includes the modal participating mass ratio for the structure Uz degree of freedom. The ratio applies to the specified mode.

SumUx

This is an array that includes the cumulative sum of the modal participating mass ratios for the structure Ux degree of freedom.

SumUy

This is an array that includes the cumulative sum of the modal participating mass ratios for the structure Uy degree of freedom.

SumUz

This is an array that includes the cumulative sum of the modal participating mass ratios for the structure Uz degree of freedom.

Rx

This is an array that includes the modal participating mass ratio for the structure Rx degree of freedom. The ratio applies to the specified mode.

Ry

This is an array that includes the modal participating mass ratio for the structure Ry degree of freedom. The ratio applies to the specified mode.

Rz

This is an array that includes the modal participating mass ratio for the structure Rz degree of freedom. The ratio applies to the specified mode.

SumRx

This is an array that includes the cumulative sum of the modal participating mass ratios for the structure Rx degree of freedom.

SumRy

This is an array that includes the cumulative sum of the modal participating mass ratios for the structure Ry degree of freedom.

SumRz

This is an array that includes the cumulative sum of the modal participating mass ratios for the structure Rz degree of freedom.

## Remarks

This function reports the modal participating mass ratios for each mode of each selected modal analysis case.

The function returns zero if the data is successfully recovered; otherwise it returns a nonzero value.

See [Analysis Results Remarks](Analysis_Results_Remarks.md) for more information.

## VBA Example



```
Sub GetModalParticipatingMassRatios()

       'dimension variables

          Dim SapObject as cOAPI

          Dim SapModel As cSapModel

          Dim ret As Long

          Dim NumberResults As Long

          Dim LoadCase() As String

          Dim StepType() As String

          Dim StepNum() As Double

          Dim Period() As Double

          Dim Ux() As Double

            Dim Uy() As Double

            Dim Uz() As Double

            Dim SumUx() As Double

            Dim SumUy() As Double

            Dim SumUz() As Double

            Dim Rx() As Double

            Dim Ry() As Double

            Dim Rz() As Double

            Dim SumRx() As Double

          Dim SumRy() As Double

          Dim SumRz() As Double



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



       'get modal participating mass ratios

          ret = SapModel.Results.ModalParticipatingMassRatios(NumberResults, LoadCase, StepType, StepNum, Period, Ux, Uy, Uz, SumUx, SumUy, SumUz, Rx, Ry, Rz, SumRx, SumRy, SumRz)



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

[ModalParticipationFactors](ModalParticipationFactors.md)

[ModalLoadParticipationRatios](ModalLoadParticipationRatios.md)

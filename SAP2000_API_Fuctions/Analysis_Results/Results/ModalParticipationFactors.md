# ModalParticipationFactors

## Syntax

SapObject.SapModel.Results.ModalParticipationFactors

## VB6 Procedure

Function ModalParticipationFactors(ByRef NumberResults As Long, ByRef LoadCase() As String, ByRef StepType() As String, ByRef StepNum() As Double, ByRef Period() As Double, ByRef Ux() As Double, ByRef Uy() As Double, ByRef Uz() As Double, ByRef Rx() As Double, ByRef Ry() As Double, ByRef Rz() As Double, ByRef ModalMass() As Double, ByRef ModalStiff() As Double) As Long

## Parameters

NumberResults

The total number of results returned by the program.

LoadCase

This is an array that includes the name of the modal load case associated with each result.

StepType

This is an array that includes the step type, if any, for each result. For modal results, this will always be Mode.See [Analysis Results Step Type](Analysis_Results_Step_Type.md).

StepNum

This is an array that includes the step number for each result. For modal results, this will always be the mode number.  See [Analysis Results Step Number](Analysis_Results_Step_Number.md).

Period

This is an array that includes the period for each result. [s]

Ux

This is an array that includes the modal participation factor for the structure Ux degree of freedom. The factor applies to the specified mode. [Fs2]

Uy

This is an array that includes the modal participation factor for the structure Uy degree of freedom. The factor applies to the specified mode. [Fs2]

Uz

This is an array that includes the modal participation factor for the structure Uz degree of freedom. The factor applies to the specified mode. [Fs2]

Rx

This is an array that includes the modal participation factor for the structure Rx degree of freedom. The factor applies to the specified mode. [FLs2]

Ry

This is an array that includes the modal participation factor for the structure Ry degree of freedom. The factor applies to the specified mode. [FLs2]

Rz

This is an array that includes the modal participation factor for the structure Rz degree of freedom. The factor applies to the specified mode. [FLs2]

ModalMass

This is an array that includes the modal mass for the specified mode.  This is a measure of the kinetic energy in the structure as it is deforming in the specified mode. [FLs2]

ModalStiff

This is an array that includes the modal stiffness for the specified mode.  This is a measure of the strain energy in the structure as it is deforming in the specified mode. [FL]

## Remarks

This function reports the modal participation factors for each mode of each selected modal analysis case.

The function returns zero if the data is successfully recovered; otherwise it returns a nonzero value.

See [Analysis Results Remarks](Analysis_Results_Remarks.md) for more information.

## VBA Example



```
Sub GetModalParticipationFactors()

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

            Dim Rx() As Double

            Dim Ry() As Double

            Dim Rz() As Double

            Dim ModalMass() As Double

          Dim ModalStiff() As Double



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



       'get modal participation factors

          ret = SapModel.Results.ModalParticipationFactors(NumberResults, LoadCase, StepType, StepNum, Period, Ux, Uy, Uz, Rx, Ry, Rz, ModalMass, ModalStiff)



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

[ModalLoadParticipationRatios](ModalLoadParticipationRatios.md)

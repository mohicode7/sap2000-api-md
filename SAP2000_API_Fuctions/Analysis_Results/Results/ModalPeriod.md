# ModalPeriod

## Syntax

SapObject.SapModel.Results.ModalPeriod

## VB6 Procedure

Function ModalPeriod(ByRef NumberResults As Long, ByRef LoadCase() As String, ByRef StepType() As String, ByRef StepNum() As Double, ByRef Period() As Double, ByRef Frequency() As Double, ByRef CircFreq() As Double, ByRef EigenValue() As Double) As Long

## Parameters

NumberResults

The number total of results returned by the program.

LoadCase

This is an array that includes the name of the modal analysis case associated with each result.

StepType

This is an array that includes the step type, if any, for each result. For modal results this is always be Mode. See [Analysis Results Step Type](Analysis_Results_Step_Type.md).

StepNum

This is an array that includes the step number for each result. For modal results this is always the mode number. See [Analysis Results Step Number](Analysis_Results_Step_Number.md).

Period

This is an array that includes the period for each result. [s]

Frequency

This is an array that includes the cyclic frequency for each result. [1/s]

CircFreq

This is an array that includes the circular frequency for each result. [rad/s]

EigenValue

This is an array that includes the eigenvalue for the specified mode for each result. [rad2/s2]

## Remarks

This function reports the modal period, cyclic frequency, circular frequency and eigenvalue for each selected modal load case.

The function returns zero if the data is successfully recovered, otherwise it returns a nonzero value.

See [Analysis Results Remarks](Analysis_Results_Remarks.md) for more information.

## VBA Example



```
Sub GetPeriod()

       'dimension variables

          Dim SapObject as cOAPI

          Dim SapModel As cSapModel

          Dim ret As Long

          Dim NumberResults As Long

          Dim LoadCase() As String

          Dim StepType() As String

          Dim StepNum() As Double

          Dim Period() As Double

          Dim Frequency() As Double

          Dim CircFreq() As Double

          Dim EigenValue() As Double



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



       'get modal period

          ret = SapModel.Results.ModalPeriod(NumberResults, LoadCase, StepType, StepNum, Period, Frequency, CircFreq, EigenValue)



       'close Sap2000

          SapObject.ApplicationExit False

          Set SapModel = Nothing

          Set SapObject = Nothing

    End Sub
  
```



## Release Notes

Initial release in version 11.00.

Changed nomenclature from Load Cases, Analysis Cases and Response Combinations to Load Patterns, Load Cases and Load Combinations, respectively, in version 12.00.

## See Also

[ModeShape](ModeShape.md)

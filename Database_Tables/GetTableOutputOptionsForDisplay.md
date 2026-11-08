# GetTableOutputOptionsForDisplay

## Syntax

SapObject.SapModel.DatabaseTables.GetTableOutputOptionsForDisplay

## VB6 Procedure

Function GetTableOutputOptionsForDisplay (ByRef BaseReactionGX
As Double, ByRef BaseReactionGY As Double, ByRef BaseReactionGZ As Double,
ByRef IsAllModes As Boolean, ByRef StartMode As Integer, ByRef EndMode
As Integer, ByRef IsAllBucklingModes As Boolean, ByRef StartBucklingMode
As Integer, ByRef EndBucklingMode As Integer, ByRef ModalHistory As Integer,
ByRef DirectHistory As Integer, ByRef NonlinearStatic As Integer, ByRef
MultistepStatic As Integer, ByRef SteadyState As Integer, ByRef SteadyStateOption
As Integer, ByRef PowerSpectralDensity As Integer, ByRef Combo As Integer,
ByRef BridgeDesign As Integer) As Integer

## Parameters

BaseReactionGX

The global X coordinate of the base reaction location.

BaseReactionGY

The global Y coordinate of the base reaction location.

BaseReactionGZ

The global Z coordinate of the base reaction location.

IsAllModes

This item is true if results are displayed for all modes
of modal load cases.

StartMode

The first mode for which results are shown for modal
load cases. This item only applies when the IsAllModes item is False.

EndMode

The last mode for which results are shown for modal
load cases. This item only applies when the IsAllModes item is False.

IsAllBucklingModes

This item is true if results are displayed for all modes
of buckling load cases.

StartBucklingMode

The first mode for which results are shown for buckling
load cases. This item only applies when the IsAllBucklingModes item is
False.

EndBucklingMode

The last mode for which results are shown for buckling
load cases. This item only applies when the IsAllBucklingModes item is
False.

ModalHistory

Indicates how multistep modal time history load case
results are displayed: 1=Envelopes, 2=Step-by-step, 3=Last Step.

DirectHistory

Indicates how direct integration time history load case
results are displayed: 1=Envelopes, 2=Step-by-step, 3=Last Step.

NonlinearStatic

Indicates how nonlinear static load case results are
displayed: 1=Envelopes, 2=Step-by-step, 3=Last Step.

MultistepStaticStatic

Indicates how multistep static load case results are
displayed: 1=Envelopes, 2=Step-by-step, 3=Last Step.

SteadyState

Indicates how steady state load case results are displayed:
1=Envelopes, 2=At Frequencies.

SteadyStateOption

Indicates how type of steady state load case results
displayed: 1=In and Out of Phase, 2=Magnitude, 3=All.

PowerSpectralDensity

Indicates how power spectral density load case results
are displayed: 1=RMS, 2=sqrt(PSD).

Combo

Indicates how load combination results are displayed:
1=Envelopes, 2=Multiple Values If Possible, 3=Correspondence.

BridgeDesign

Indicates which bridge design results are displayed:
1=Controlling Combo, 2=All Combos.

## Remarks

This function returns the output option data used for
display of analysis result tables.

The function returns zero if the information is successfully
retrieved, otherwise it returns a nonzero value.

## VBA Example

This example assumes that a file MyModel.sdb exists.



```
Sub GetTableOutputOptions()

   'dimension variables

      Dim SapObject as cOAPI

      Dim SapModel As cSapModel

      Dim ret As Long

      Dim FileName As String

      Dim BaseReactionGX As Double

      Dim BaseReactionGY As Double

      Dim BaseReactionGZ As Double

      Dim IsAllModes As Boolean

      Dim StartMode As Integer

      Dim EndMode As Integer

      Dim IsAllBucklingModes As Boolean

      Dim StartBucklingMode As Integer

      Dim EndBucklingMode As Integer

      Dim ModalHistory As Integer

      Dim DirectHistory As Integer

      Dim NonlinearStatic As Integer

      Dim MultistepStatic As Integer

      Dim SteadyState As Integer

      Dim SteadyStateOption As Integer

      Dim PowerSpectralDensity As Integer

      Dim Combo As Integer

      Dim BridgeDesign As Integer



   'create Sap2000 object

      Set SapObject = CreateObject("CSI.SAP2000.API.SapObject")



   'start Sap2000 application

      SapObject.ApplicationStart



   'create SapModel object

      Set SapModel = SapObject.SapModel



   'initialize model

      ret = SapModel.InitializeNewModel



   'open an existing file

      FileName = "C:\SapAPI\MyModel.sdb"

      ret = SapModel.File.OpenFile(FileName)



   'get table output options used for display in the analysis 
 result tables

      ret = SapModel. DatabaseTables.GetTableOutputOptionsForDisplay 
 (BaseReactionGX, BaseReactionGY, BaseReactionGZ, IsAllModes, StartMode, 
 EndMode, IsAllBucklingModes,
StartBucklingMode, EndBucklingMode, ModalHistory, DirectHistory, 
 NonlinearStatic, MultistepStatic, SteadyState, SteadyStateOption, PowerSpectralDensity, 
 Combo, BridgeDesign)



   'close Sap2000

      SapObject.ApplicationExit False

      Set SapModel = Nothing

      Set SapObject = Nothing

End Sub
```



## Release Notes

Initial release in version 23.00.

## See Also

[SetTableOutputOptionsForDisplay](SetTableOutputOptionsForDisplay.md)

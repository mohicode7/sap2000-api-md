# SetTableOutputOptionsForDisplay

## Syntax

SapObject.SapModel.DatabaseTables.SetTableOutputOptionsForDisplay

## VB6 Procedure

Function SetTableOutputOptionsForDisplay (ByVal BaseReactionGX
As Double, ByVal BaseReactionGY As Double, ByVal BaseReactionGZ As Double,
ByVal IsAllModes As Boolean, ByVal StartMode As Integer, ByVal EndMode
As Integer, ByVal IsAllBucklingModes As Boolean, ByVal StartBucklingMode
As Integer, ByVal EndBucklingMode As Integer, ByVal ModalHistory As Integer,
ByVal DirectHistory As Integer, ByVal NonlinearStatic As Integer, ByVal
MultistepStatic As Integer, ByVal SteadyState As Integer, ByVal SteadyStateOption
As Integer, ByVal PowerSpectralDensity As Integer, ByVal Combo As Integer,
ByVal BridgeDesign As Integer) As Integer

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

**EndMode**

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

This function sets the output option data used for display
of analysis result tables.

The function returns zero if the information is successfully
set, otherwise it returns a nonzero value.

## VBA Example

This example assumes that a file MyModel.sdb exists.



```
Sub SetTableOutputOptions()

   'dimension variables

      Dim SapObject as cOAPI

      Dim SapModel As cSapModel

      Dim ret As Long

      Dim FileName As String



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



   'set table output options used for display in the analysis 
 result tables

      ret = SapModel. DatabaseTables.SetTableOutputOptionsForDisplay 
 (0, 0, 0, False, 1, 12, False, 1, 6, 1, 1, 2, 2, 1, 1, 2, 2, 1)



   'close Sap2000

      SapObject.ApplicationExit False

      Set SapModel = Nothing

      Set SapObject = Nothing

End Sub
```



## Release Notes

Initial release in version 23.00.

## See Also

[GetTableOutputOptionsForDisplay](GetTableOutputOptionsForDisplay.md)

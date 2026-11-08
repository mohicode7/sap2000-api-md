# SetFrameForceMultiple

## Syntax

SapObject.SapModel.ExternalAnalysisResults.SetFrameForceMultiple

## VB6 Procedure

Function SetFrameForceMultiple(ByVal NumberFrameNames As Long, ByVal FrameName() As String, ByVal NumberLoadCases As Long, ByVal LoadCase() As String, ByVal FirstStep() As Long, ByVal LastStep() As Long, ByRef P() As Double, ByRef M2() As Double, ByRef M3() As Double) As Long

## Parameters

NumberFrameNames

The number of input frames

FrameName

An array of existing frame objects

NumberLoadCases

The number of input load cases

LoadCase

An array of existing external results load cases for which user-supplied external analysis results relevant to the frame objects are available.

FirstStep

An array of the first step values for each load case. This should be of length NumberLoadCases.

LastStep

An array of the last step values for each load case. This should be of length NumberLoadCases.

P

An array of the axial forces for each frame station. [F]

V2

An array of the shear forces in the local 2 direction for each frame station. [F]

V3

An array of the shear forces in the local 3 direction for each frame station. [F]

T

An array of the torsion for each frame station. [F]

M2

An array of the moment about the local 2-axis for each frame station. [F]

M3

An array of the moment about the local 3-axis for each frame station. [F]

## Remarks

This function sets the external result forces for all input frames and load cases.

The P, V2, V3, T, M2, M3 arrays should be of length:

(Sum of all steps for all input Load Cases) \* (Sum of all stations for all input Frames)

As an example, consider setting results for two Frames, each with two defined Stations, for two Load Cases, each Load Case containing three steps, eg First Step = 1, Last step = 3.

The values of this array will be iterated over in the following order:

The P[0] value will be the axial force for FrameName[0], LoadCase[0], FirstStep[0], 1st Station

The P[1] value will be the axial force for FrameName[0], LoadCase[0], FirstStep[0], 2nd Station

The P[2] value will be the axial force for FrameName[0], LoadCase[0], FirstStep[0]+1, 1st Station

The P[3] value will be the axial force for FrameName[0], LoadCase[0], FirstStep[0]+1, 2nd Station

The P[4] value will be the axial force for FrameName[0], LoadCase[0], LastStep[0], 1st Station

The P[5] value will be the axial force for FrameName[0], LoadCase[0], LastStep[0], 2nd Station

The P[6] value will be the axial force for FrameName[0], LoadCase[1], FirstStep[1], 1st Station

The P[7] value will be the axial force for FrameName[0], LoadCase[1], FirstStep[1], 2nd Station

The P[8] value will be the axial force for FrameName[0], LoadCase[1], FirstStep[1]+1, 1st Station

The P[9] value will be the axial force for FrameName[0], LoadCase[1], FirstStep[1]+1, 2nd Station

The P[10] value will be the axial force for FrameName[0], LoadCase[1], LastStep[1], 1st Station

The P[11] value will be the axial force for FrameName[0], LoadCase[1],  LastStep[1], 2nd Station

The P[12] value will be the axial force for FrameName[1], LoadCase[0], FirstStep[0], 1st Station

And so on…

The number of stations must be previously declared using SetFrameStations

Enter 0 for any unneeded values in these arrays

The function returns zero if the forces are successfully set, otherwise it returns a nonzero value.

## VBA Example

## Release Notes

Initial Release in version 16.0.0.

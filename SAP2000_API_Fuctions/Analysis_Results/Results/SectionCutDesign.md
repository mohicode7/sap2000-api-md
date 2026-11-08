# SectionCutDesign

## Syntax

SapObject.SapModel.Results.SectionCutDesign

## VB6 Procedure

Function SectionCutDesign(ByRef NumberResults As Long, ByRef SCut() As String, ByRef LoadCase() As String, ByRef StepType() As String, ByRef StepNum() As Double, ByRef P() As Double, ByRef V2() As Double, ByRef V3() As Double, ByRef T() As Double, ByRef M2() As Double, ByRef M3() As Double) As Long

## Parameters

NumberResults

The total number of results returned by the program.

SCut

This is an array that includes the name of the section cut associated with each result.

LoadCase

This is an array that includes the name of the analysis case or load combination associated with each result.

StepType

This is an array that includes the step type, if any, for each result. See [Analysis Results Step Type](Analysis_Results_Step_Type.md).

StepNum

This is an array that includes the step number, if any, for each result. See [Analysis Results Step Number](Analysis_Results_Step_Number.md).

P, V2, V3

These are one dimensional arrays that include the axial force, shear force in the section cut local 2 direction and shear force in the section cut local 3 direction, respectively, for each result. [F]

T, M2, M3

These are one dimensional arrays that include the torsion, moment about the section cut local 2 axis and moment about the section cut local 3-axis, respectively, for each result. [FL]

## Remarks

This function reports the section cut force for sections cuts that are specified to have a Design (P, V2, V3, T, M2, M3) result type.

The function returns zero if the section cut forces are successfully recovered, otherwise it returns a nonzero value.

See [Analysis Results Remarks](Analysis_Results_Remarks.md) for more information.

## VBA Example



```
Similar to FrameForce



       'get section cut forces with design output convention

          ret = SapModel.Results.SectionCutDesign(NumberResults, SCut, LoadCase, StepType, StepNum, P, V2, V3, T, M2, M3)
  
```



## Release Notes

Initial release in version 11.00.

Changed nomenclature from Load Cases, Analysis Cases and Response Combinations to Load Patterns, Load Cases and Load Combinations, respectively, in version 12.00.

## See Also

[SectionCutAnalysis](SectionCutAnalysis.md)

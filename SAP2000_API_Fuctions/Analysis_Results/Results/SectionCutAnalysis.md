# SectionCutAnalysis

## Syntax

SapObject.SapModel.Results.SectionCutAnalysis

## VB6 Procedure

Function SectionCutAnalysis(ByRef NumberResults As Long, ByRef SCut() As String, ByRef LoadCase() As String, ByRef StepType() As String, ByRef StepNum() As Double, ByRef F1() As Double, ByRef F2() As Double, ByRef F3() As Double, ByRef M1() As Double, ByRef M2() As Double, ByRef M3() As Double) As Long

## Parameters

NumberResults

The number total of results returned by the program.

SCut

This is an array that includes the name of the section cut associated with each result.

LoadCase

This is an array that includes the name of the analysis case or load combination associated with each result.

StepType

This is an array that includes the step type, if any, for each result. See [Analysis Results Step Type](Analysis_Results_Step_Type.md).

StepNum

This is an array that includes the step number, if any, for each result. See [Analysis Results Step Number](Analysis_Results_Step_Number.md).

F1, F2, F3

These are one dimensional arrays that include the forces in the section cut local axes directions for each result. [F]

M1, M2, M3

These are one dimensional arrays that include the moments about the section cut local axes for each result. [FL]

## Remarks

This function reports the section cut force for sections cuts that are specified to have an Analysis (F1, F2, F3, M1, M2, M3) result type.

The function returns zero if the section cut forces are successfully recovered, otherwise it returns a nonzero value.

See [Analysis Results Remarks](Analysis_Results_Remarks.md) for more information.

## VBA Example



```
Similar to FrameJointForce



       'get section cut forces with analysis output convention

          ret = SapModel.Results.SectionCutAnalysis(NumberResults, SCut, LoadCase, StepType, StepNum, F1, F2, F3, M1, M2, M3)
  
```



## Release Notes

Initial release in version 11.00.

Changed nomenclature from Load Cases, Analysis Cases and Response Combinations to Load Patterns, Load Cases and Load Combinations, respectively, in version 12.00.

## See Also

[SectionCutDesign](SectionCutDesign.md)

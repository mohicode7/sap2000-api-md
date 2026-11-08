# Analysis Results Remarks

The analysis results are returned in a collection of one-dimensional result arrays. Each result array is created as a dynamic array by the API user. As an example, in VBA a dynamic string array is defined by:



```
   Dim MyArray() as String
```



The array is dimensioned to (NumberResults– 1) inside the Sap2000 program, filled with the result values, and returned to the API user.

The arrays are zero-based. Thus the first result is at array index 0, and the last result is at array index (NumberResults- 1). For example, the StepType() array is filled as:



```
   StepType(0) = Step type for first result returned

   StepType(1) = Step type for second result

               .

               .

   StepType(NumberResults- 1) = Step type for last result
```



Immediately before requesting results data, it is a good idea to clear the SelectedForOutput flag for all load cases and response combinations and then to set the flag True for those cases and combos for which output is to be generated. This avoids confusion as to which cases and combos are currently selected for output.

## Release Notes

Changed nomenclature from Load Cases, Analysis Cases and Response Combinations to Load Patterns, Load Cases and Load Combinations, respectively, in version 12.00.

# GetSolverOption\_3

## Syntax

SapObject.SapModel.Analyze.GetSolverOption\_3

## VB6 Procedure

Function GetSolverOption\_3(ByRef SolverType As Long, ByRef SolverProcessType As Long, ByRef NumberParallelRuns As Long, ByRef ResponseFileSizeMaxMB As Long, ByRef NumberAnalysisThreads As Long, ByRef StiffCase As String) As Long

## Parameters

**SolverType**

This is 0, 1 or 2, indicating the solver type.

0 = Standard solver  
1 = Advanced solver  
2 = Multi-threaded solver

**SolverProcessType**

This is 0, 1 or 2, indicating the process the analysis is run.

0 = Auto (program determined)  
1 = GUI process  
2 = Separate process

**NumberParallelRuns**

This is an integer between -8 and 8, inclusive, not including -1 or 0.

-8 to -2 = The negative of the program determined value when the assigned value is 0 = Auto parallel (use up to all physical cores - max 8).  
1 = Serial.  
2 to 8 = User defined parallel (use up to this fixed number of cores - max 8

ResponseFileSizeMaxMB

The maximum size of a response file in MB before a new response file is created. Positive if user specified, negative if program determined.

NumberAnalysisThreads

Number of threads that the analysis can use. Positive if user specified, negative if program determined.

**StiffCase**

The name of the load case used when outputting the mass and stiffness matrices to text files. If this item is blank, no matrices are output.

## Remarks

This function retrieves the model solver options.

The function returns zero if the options are successfully retrieved; otherwise it returns a nonzero value.

## **VBA Example**



```
Sub GetModelSolverOption()

       'dimension variables

          Dim SapObject as cOAPI

          Dim SapModel As cSapModel

          Dim ret As Long

          Dim SolverType As Long

          Dim SolverProcessType As Long

          Dim StiffCase As String



       'create Sap2000 object

          Set SapObject = CreateObject("CSI.SAP2000.API.SapObject")



       'start Sap2000 application

          SapObject.ApplicationStart



       'create SapModel object

          Set SapModel = SapObject.SapModel



       'initialize model

          ret = SapModel.InitializeNewModel



       'create model from template

          ret = SapModel.File.New2DFrame(PortalFrame, 2, 144, 2, 288)



       'set model solver options

          ret = SapModel.Analyze.GetSolverOption_3(SolverType, SolverProcessType, NumberParallelRuns, ResponseFileSizeMaxMB, NumberAnalysisThreads, StiffCase)



       'close Sap2000

          SapObject.ApplicationExit False

          Set SapModel = Nothing

          Set SapObject = Nothing

    End Sub
  
```



## **Release Notes**

Initial release in v21.1.0.

Modified the allowed values for NumberParallelRuns in v 22.1.0.

Added ResponseFileSizeMaxMB and NumberAnalysisThreads parameters in v23.2.0

This function supersedes [GetSolverOption\_2](../Obsolete_Functions/NEW_-_Analyze.GetSolverOption_2.md), adding the NumberParallelRuns parameter and removing the previous Force32BitSolver parameter

## See Also

[SetSolverOption\_3](Analyze.SetSolverOption_3.md)

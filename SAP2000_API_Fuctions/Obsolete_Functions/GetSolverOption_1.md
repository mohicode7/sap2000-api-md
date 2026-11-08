# GetSolverOption\_1 (Note:  Newer function available)

## Syntax

SapObject.SapModel.Analyze.GetSolverOption

## VB6 Procedure

Function GetSolverOption(ByRef SolverType As Long, ByRef
SolverProcessType As Long , ByRef Force32BitSolver As Boolean, ByRef StiffCase
As String) As Long

## Parameters

SolverType

This is 0, 1 or 2, indicating the solver type.

0 = Standard solver

1 = Advanced solver

2
= Multi-threaded solver

SolverProcessType

This is 0, 1 or 2, indicating the process the analysis
is run.

0 = Auto (program determined)

1 = GUI process

2
= Separate process

Force32BitSolver

This is True if the analysis is always run using 32-bit,
even on 64-bit computers.

StiffCase

The name of the load case used when outputting the mass
and stiffness matrices to text files. If this item is blank, no matrices
are output.

## Remarks

This function retrieves the model solver options.

The function returns zero if the options are successfully
retrieved; otherwise it returns a nonzero value.

## VBA Example



```
Sub GetModelSolverOption()

   'dimension variables

      Dim SapObject as cOAPI

      Dim SapModel As cSapModel

      Dim ret As Long

      Dim SolverType As Long

      Dim SolverProcessType As Long

      Dim Force32BitSolver As Boolean

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

      ret = SapModel.File.New2DFrame(PortalFrame, 
 2, 144, 2, 288)



   'set model solver options

      ret = SapModel.Analyze.GetSolverOption_1(SolverType, 
 SolverProcessType , Force32BitSolver, StiffCase)



   'close Sap2000

      SapObject.ApplicationExit False

      Set SapModel = Nothing

      Set SapObject = Nothing

End Sub
```



## Release Notes

Initial release in version 14.2.2.

This function superseded [GetSolverOption.](GetSolverOption.md)

This function is obsolete and has been superseded by
[GetSolverOption\_2](NEW_-_Analyze.GetSolverOption_2.md)
as of version 21.1.0. This function is maintained for backwards compatibility.

## See Also

[SetSolverOption\_1](SetSolverOption_1.md)

# SetSolverOption\_1 (Note:  Newer function available)

## Syntax

SapObject.SapModel.Analyze.SetSolverOption\_1

## VB6 Procedure

Function SetSolverOption(ByVal SolverType As Long, ByVal
SolverProcessType As Long, ByVal Force32BitSolver As Boolean, Optional
ByVal StiffCase As String = "") As Long

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
and stiffness matrices to text files If this item is blank, no matrices
are output.

## Remarks

This function sets the model solver options.

The function returns zero if the options are successfully
set; otherwise it returns a nonzero value.

## VBA Example



```
Sub SetModelSolverOption()

   'dimension variables

      Dim SapObject as 
 cOAPI

      Dim SapModel As cSapModel

      Dim ret As Long



   'create Sap2000 object

      Set SapObject = 
 New SAP2000.SapObject



   'start Sap2000 application

      SapObject.ApplicationStart



   'create SapModel object

      Set SapModel = 
 SapObject.SapModel



   'initialize model

      ret = SapModel.InitializeNewModel



   'create model from template

      ret = SapModel.File.New2DFrame(PortalFrame, 
 2, 144, 2, 288)



   'set model solver options

      ret = SapModel.Analyze.SetSolverOption_1(1, 
 1, True, "DEAD")



   'close Sap2000

      SapObject.ApplicationExit 
 False

      Set SapModel = Nothing

      Set SapObject = Nothing

End Sub
```



## Release Notes

Initial release in version 14.2.2

This function is obsolete and has been superseded by
[SetSolverOption\_2](NEW_-_Analyze.SetSolverOption_2.md) as
of version 21.1.0. This function is maintained for backwards compatibility.

This function supersedes [SetSolverOption](SetSolverOption.md).

## See Also

[GetSolverOption\_1](GetSolverOption_1.md)

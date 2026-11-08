# SetSolverOption (Note:  Newer function available)

## Syntax

SapObject.SapModel.Analyze.SetSolverOption

## VB6 Procedure

Function SetSolverOption(ByVal SolverType As Long, ByVal Force32BitSolver As Boolean, Optional ByVal StiffCase As String = "") As Long

## Parameters

SolverType

This is 0 or 1, indicating the solver type.

0 = Standard solver

1 = Advanced solver

Force32BitSolver

This is True if the analysis is always run using 32-bit, even on 64-bit computers.

StiffCase

The name of the load case used when outputting the mass and stiffness matrices to text files. If this item is blank, no matrices are output.

## Remarks

This function sets the model solver options.

The function returns zero if the options are successfully set; otherwise it returns a nonzero value.

## VBA Example



```
Sub SetModelSolverOption()

   'dimension variables

      Dim SapObject As Sap2000v15.SapObject

      Dim SapModel As cSapModel

      Dim ret As Long



   'create Sap2000 object

      Set SapObject = New Sap2000v15.SapObject



   'start Sap2000 application

      SapObject.ApplicationStart



   'create SapModel object

      Set SapModel = SapObject.SapModel



   'initialize model

      ret = SapModel.InitializeNewModel



   'create model from template

      ret = SapModel.File.New2DFrame(PortalFrame, 2, 144, 2, 288)



   'set model solver options

      ret = SapModel.Analyze.SetSolverOption(1, True, "DEAD")



   'close Sap2000

      SapObject.ApplicationExit False

      Set SapModel = Nothing

      Set SapObject = Nothing

End Sub
```



## Release Notes

Initial release in version 11.03.

Changed nomenclature from Load Cases, Analysis Cases and Response Combinations to Load Patterns, Load Cases and Load Combinations, respectively, in version 12.00.

This function is obsolete and has been superseded by [SetSolverOption\_1](SetSolverOption_1.md) as of version 14.2.2. This function is maintained for backwards compatibility.

## See Also

[GetSolverOption](GetSolverOption.md)

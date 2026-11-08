# **SetSolverOption\_2** (Note:  Newer function available)

## **Syntax**

SapObject.SapModel.Analyze.SetSolverOption\_2

## **VB6 Procedure**

Function SetSolverOption\_2(ByVal SolverType As Long,
ByVal SolverProcessType As Long, ByVal NumberParallelRuns As Long, Optional
ByVal StiffCase As String = "") As Long

## **Parameters**

**SolverType**

This is 0, 1 or 2, indicating the solver type.

0 = Standard
solver  
1 = Advanced solver  
2 = Multi-threaded solver

**SolverProcessType**

This is 0, 1 or 2, indicating the process the analysis
is run.

0 = Auto
(program determined)  
1 = GUI process  
2 = Separate process

**NumberParallelRuns**

This is an integer between -8 and 8, inclusive, not
including -1.

            -8
to -2 = Auto parallel (use up to all physical cores - max 8). Treated
the same as 0.  
            -1
= Illegal value; will return an error.  
            0
= Auto parallel (use up to all physical cores).  
            1
= Serial.  
            2
to 8 = User defined parallel (use up to this fixed number of cores -
max 8).

**StiffCase**

The name of the load case used when outputting the mass
and stiffness matrices to text files If this item is blank, no matrices
are output.

## **Remarks**

This function sets the model solver options.

The function returns zero if the options are successfully
set; otherwise it returns a nonzero value.

## **VBA Example**



```
Sub SetModelSolverOption()

   'dimension variables

      Dim SapObject as cOAPI

      Dim SapModel As cSapModel

      Dim ret As Long



   'create Sap2000 object

      Set SapObject = New SAP2000.SapObject



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

      ret = SapModel.Analyze.SetSolverOption_2(1, 
 1, 3, "DEAD")



   'close Sap2000

      SapObject.ApplicationExit False

      Set SapModel = Nothing

      Set SapObject = Nothing

End Sub
```



## **Release Notes**

Initial release in v21.1.0.

Modified the allowed values for NumberParallelRuns in
v 22.1.0.

This function is obsolete and has been superseded by
[SetSolverOption\_3](../Analyze/Analyze.SetSolverOption_3.md)
as of version 23.2.0. This function is maintained for backwards compatibility.

This function supersedes [SetSolverOption\_1](SetSolverOption_1.md),
adding the NumberParallelRuns parameter and removing the previous Force32BitSolver
parameter.

## **See Also**

[GetSolverOption\_2](NEW_-_Analyze.GetSolverOption_2.md)

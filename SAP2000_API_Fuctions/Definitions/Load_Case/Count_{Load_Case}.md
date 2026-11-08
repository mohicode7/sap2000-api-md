# Count

## Syntax

SapObject.SapModel.LoadCases.Count

## VB6 Procedure

Function Count(Optional ByVal CaseType As eLoadCaseType) As Long

## Parameters

CaseType

This optional value is one of the following items in the eLoadCaseType enumeration.

CASE\_LINEAR\_STATIC = 1

CASE\_NONLINEAR\_STATIC = 2

CASE\_MODAL = 3

CASE\_RESPONSE\_SPECTRUM = 4

CASE\_LINEAR\_HISTORY = 5  (Modal Time History)

CASE\_NONLINEAR\_HISTORY = 6  (Modal Time History)

CASE\_LINEAR\_DYNAMIC = 7  (Direct Integration Time History)

CASE\_NONLINEAR\_DYNAMIC = 8  (Direct Integration Time History)

CASE\_MOVING\_LOAD = 9

CASE\_BUCKLING = 10

CASE\_STEADY\_STATE = 11

CASE\_POWER\_SPECTRAL\_DENSITY = 12

CASE\_LINEAR\_STATIC\_MULTISTEP = 13

CASE\_HYPERSTATIC = 14

If no value is input for CaseType, a count is returned for all load cases in the model regardless of type.

## Remarks

This function returns the total number of defined load cases in the model. If desired, counts can be returned for all load cases of a specified type in the model.

## VBA Example



```
Sub CountLoadCases()

   'dimension variables

      Dim SapObject as cOAPI

      Dim SapModel As cSapModel

      Dim ret As Long

      Dim Count As Long



   'create Sap2000 object

      Set SapObject = CreateObject("CSI.SAP2000.API.SapObject")



   'start Sap2000 application

      SapObject.ApplicationStart



   'create SapModel object

      Set SapModel = SapObject.SapModel



   'initialize model

      ret = SapModel.InitializeNewModel



   'create model from template

      ret = SapModel.File.New2DFrame(PortalFrame, 3, 124, 3, 200)



   'return number of load cases of all types

      Count = SapModel.LoadCases.Count



   'close Sap2000

      SapObject.ApplicationExit False

      Set SapModel = Nothing

      Set SapObject = Nothing

End Sub
```



## Release Notes

Initial release in version 11.00.

Added optional CaseType parameter in version 11.02.

Changed nomenclature from Load Cases, Analysis Cases and Response Combinations to Load Patterns, Load Cases and Load Combinations, respectively, in version 12.00.

Added one item to the eLoadCaseType enumeration in version 12.00.

## See Also

# GetTypeOAPI   (Note:  Newer Function Available)

## Syntax

SapObject.SapModel.LoadCases.GetTypeOAPI

## VB6 Procedure

Function GetTypeOAPI(ByVal Name As String, ByRef CaseType As eLoadCaseType, ByRef SubType as long) As Long

## Parameters

Name

The name of an existing load case.

CaseType

This is one of the following items in the eLoadCaseType enumeration.

CASE\_LINEAR\_STATIC = 1

CASE\_NONLINEAR\_STATIC = 2

CASE\_MODAL = 3

CASE\_RESPONSE\_SPECTRUM = 4

CASE\_LINEAR\_HISTORY = 5  (Modal Time History)

CASE\_NONLINEAR\_HISTORY = 6  (Modal Time History)

CASE\_LINEAR\_DYNAMIC = 7  (Direct Integration Time History)

CASE\_NONLINEAR\_DYNAMIC = 8  (Direct Integration Time History)

CASE\_MOVING\_LOAD = 9

CASE\_BUCKLING = 10

CASE\_STEADY\_STATE = 11

CASE\_POWER\_SPECTRAL\_DENSITY = 12

CASE\_LINEAR\_STATIC\_MULTISTEP = 13

CASE\_HYPERSTATIC = 14

SubType

This is an integer representing the load case sub type. This item only applies for certain case types.

For CASE\_NONLINEAR\_STATIC:

1 = Nonlinear

2 = Nonlinear staged construction

For CASE\_MODAL:

1 = Eigen

2 = Ritz

For CASE\_LINEAR\_HISTORY:

1 = Transient

2 = Periodic

## Remarks

This function retrieves the case type for the specified load case.

The function returns zero if the type is successfully retrieved; otherwise it returns nonzero.

## VBA Example



```
Sub GetLoadCaseType()

   'dimension variables

      Dim SapObject As Sap2000v15.SapObject

      Dim SapModel As cSapModel

      Dim ret As Long

      Dim CaseType As eLoadCaseType

      Dim SubType As Long



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



   'get load case type

      ret = SapModel.LoadCases.GetTypeOAPI("DEAD", CaseType, SubType)



   'close Sap2000

      SapObject.ApplicationExit False

      Set SapModel = Nothing

      Set SapObject = Nothing

End Sub
```



## Release Notes

Initial release in version 11.02.

Changed nomenclature from Load Cases, Analysis Cases and Response Combinations to Load Patterns, Load Cases and Load Combinations, respectively, in version 12.00.

Added one item to the eLoadCaseType enumeration in version 12.00.

This function is obsolete and has been superseded by [GetType\_1](../Definitions/Load_Case/GetType_1_%7BLoad_Case%7D.md) as of version 12.00. This function is maintained for backwards compatibility.

Changed function name to GetTypeOAPI in v17.0.0.

## See Also

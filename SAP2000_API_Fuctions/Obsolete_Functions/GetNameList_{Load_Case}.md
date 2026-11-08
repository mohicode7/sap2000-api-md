# GetNameList

## Syntax

SapObject.SapModel.LoadCases.GetNameList

## VB6 Procedure

Function GetNameList(ByRef NumberNames As Long, ByRef MyName() As String, Optional ByVal CaseType As eLoadCaseType) As Long

## Parameters

NumberNames

The number of load case names retrieved by the program.

MyName

This is a one-dimensional array of load case names. The MyName array is created as a dynamic, zero-based, array by the API user:



```
Dim MyName() as String
```



The array is dimensioned to (NumberNames - 1) inside the Sap2000 program, filled with the names, and returned to the API user.

CaseType

This optional value is one of the following items in the eLoadCaseType enumeration.

LinearStatic = 1

NonlinearStatic = 2

Modal = 3

ResponseSpectrum = 4

LinearHistory = 5  (Modal Time History)

NonlinearHistory = 6  (Modal Time History)

LinearDynamic = 7  (Direct Integration Time History)

NonlinearDynamic = 8  (Direct Integration Time History)

MovingLoad = 9

Buckling = 10

SteadyState = 11

PowerSpectralDensity = 12

LinearStaticMultistep = 13

Hyperstatic = 14

ExternalResults = 15

If no value is input for CaseType, names are returned for all load cases in the model regardless of type.

## Remarks

This function retrieves the names of all defined load cases of the specified type.

The function returns zero if the names are successfully retrieved, otherwise it returns nonzero.

## VBA Example



```
Sub GetLoadCaseNames()

   'dimension variables

      Dim SapObject as cOAPI

      Dim SapModel As cSapModel

      Dim ret As Long

      Dim NumberNames As Long

      Dim MyName() As String



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



   'get load case names

      ret = SapModel.LoadCases.GetNameList(NumberNames, MyName)



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

This function is obsolete and has been superseded by GetNameList\_1 as of v21.0.0. This function is maintained for backwards compatibility where staged construction and nonlinear multi-step cases are a subtype of load case type nonlinear static.

## See Also

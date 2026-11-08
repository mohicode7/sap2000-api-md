# GetNameList\_1

## Syntax

SapObject.SapModel.LoadCases.GetNameList\_1

## VB6 Procedure

Function GetNameList\_1(ByRef NumberNames As Long, ByRef MyName() As String, Optional ByVal CaseType As eLoadCaseType) As Long

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

Linear Static = 1

Nonlinear\_Static = 2

Modal = 3

Response\_Spectrum = 4

Linear\_History = 5  (Modal Time History)

Nonlinear\_History = 6  (Modal Time History)

Linear\_Dynamic = 7  (Direct Integration Time History)

Nonlinear\_Dynamic = 8  (Direct Integration Time History)

Moving\_Load = 9

Buckling = 10

Steady\_State = 11

Power\_Spectral\_Density = 12

Linear\_Static\_Multistep = 13

Hyperstatic = 14

External\_Results = 15

StagedConstruction = 16

NonlinearStaticMultiStep = 17

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

      ret = SapModel.LoadCases.GetNameList_1(NumberNames, MyName)



   'close Sap2000

      SapObject.ApplicationExit False

      Set SapModel = Nothing

      Set SapObject = Nothing

End Sub
```



## Release Notes

Initial release in v21.0.0.

This function supersedes [GetNameList {Load Case}](../../Obsolete_Functions/GetNameList_%7BLoad_Case%7D.md).

## See Also

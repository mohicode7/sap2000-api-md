# GetNameList

## Syntax

SapObject.SapModel.Func.GetNameList

## VB6 Procedure

Function GetNameList(ByRef NumberNames As Long, ByRef MyName() As String, Optional ByVal FuncType As Long = 0) As Long

## Parameters

NumberNames

The number of function names retrieved by the program.

MyName

This is a one-dimensional array of function names. The MyName array is created as a dynamic, zero-based array by the API user:



```
Dim MyName() as String
```



The array is dimensioned to (NumberNames - 1) inside the SAP2000 program, filled with the names, and returned to the API user.

FuncType

This is one of the following numbers, indicating the type of function for which the name list is desired.

0 = All function types

1 = Response spectrum

2 = Time history

3 = Power spectral density

4 = Steady state

## Remarks

This function retrieves the names of all defined functions of the specified type.

The function returns zero if the names are successfully retrieved; otherwise it returns nonzero.

## VBA Example



```
Sub GetFunctionNames()

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



   'get function names

      ret = SapModel.Func.GetNameList(NumberNames, MyName)



   'close Sap2000

      SapObject.ApplicationExit False

      Set SapModel = Nothing

      Set SapObject = Nothing

End Sub
```



## Release Notes

Initial release in version 11.01.

## See Also

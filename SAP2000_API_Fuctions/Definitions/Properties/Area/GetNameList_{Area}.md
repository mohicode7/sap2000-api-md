# GetNameList

## Syntax

SapObject.SapModel.PropArea.GetNameList

## VB6 Procedure

Function GetNameList(ByRef NumberNames As Long, ByRef MyName() As String, Optional ByVal PropType As Long = 0) As Long

## Parameters

NumberNames

The number of area property names retrieved by the program.

MyName

This is a one-dimensional array of area property names. The MyName array is created as a dynamic, zero-based, array by the API user:



```
Dim MyName() as String
```



The array is dimensioned to (NumberNames - 1) inside the SAP2000 program, filled with the names, and returned to the API user.

PropType

This optional value is 0, 1, 2 or 3, indicating the type of area properties included in the name list.

0 = All

1 = Shell

2 = Plane

3 = Asolid

## Remarks

This function retrieves the names of all defined area properties of the specified type.

The function returns zero if the names are successfully retrieved; otherwise it returns nonzero.

## VBA Example



```
Sub GetAreaPropNames()

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

      ret = SapModel.File.NewWall(2, 48, 2, 48)



   'get area property names

      ret = SapModel.PropArea.GetNameList(NumberNames, MyName)



   'close Sap2000

      SapObject.ApplicationExit False

      Set SapModel = Nothing

      Set SapObject = Nothing

End Sub
```



## Release Notes

Initial release in version 11.02.

## See Also

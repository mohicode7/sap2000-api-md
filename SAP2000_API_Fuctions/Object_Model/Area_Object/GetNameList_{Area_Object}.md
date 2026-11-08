# GetNameList

## Syntax

SapObject.SapModel.AreaObj.GetNameList

## VB6 Procedure

Function GetNameList(ByRef NumberNames As Long, ByRef MyName() As String) As Long

## Parameters

NumberNames

The number of area object names retrieved by the program.

MyName

This is a one-dimensional array of area object names. The MyName array is created as a dynamic, zero-based, array by the API user:



```
Dim MyName() as String

 
```



The array is dimensioned to (NumberNames – 1) inside the Sap2000 program, filled with the names, and returned to the API user.

## Remarks

This function retrieves the names of all defined area objects.

The function returns zero if the names are successfully retrieved, otherwise it returns nonzero.

## VBA Example



```
Sub GetAreaObjectNames()

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



   'get area object names

      ret = SapModel.AreaObj.GetNameList(NumberNames, MyName)



   'close Sap2000

      SapObject.ApplicationExit False

      Set SapModel = Nothing

      Set SapObject = Nothing

End Sub
```



## Release Notes

Initial release in version 11.00.

## See Also

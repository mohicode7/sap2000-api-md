# GetNameList

## Syntax

SapObject.SapModel.SolidObj.GetNameList

## VB6 Procedure

Function GetNameList(ByRef NumberNames As Long, ByRef MyName() As String) As Long

## Parameters

NumberNames

The number of solid object names retrieved by the program.

MyName

This is a one-dimensional array of solid object names. The MyName array is created as a dynamic, zero-based, array by the APIuser:



```
Dim MyName() as String
```



The array is dimensioned to (NumberNames – 1) inside the SAP2000 program, filled with the names, and returned to the APIuser.

## Remarks

This function retrieves the names of all defined solid objects.

The function returns zero if the names are successfully retrieved; otherwise it returns nonzero.

## VBA Example



```
Sub GetSolidObjectNames()

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

      ret = SapModel.File.NewSolidBlock(300, 400, 200, , , 2, 2, 2)



   'get solid object names

      ret = SapModel.SolidObj.GetNameList(NumberNames, MyName)



   'close Sap2000

      SapObject.ApplicationExit False

      Set SapModel = Nothing

      Set SapObject = Nothing

End Sub
```



## Release Notes

Initial release in version 11.00.

## See Also

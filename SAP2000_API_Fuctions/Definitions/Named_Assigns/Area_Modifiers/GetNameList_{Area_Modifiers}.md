# GetNameList

## Syntax

SapObject.SapModel.NamedAssign.ModifierArea.GetNameList

## VB6 Procedure

Function GetNameList(ByRef NumberNames As Long, ByRef MyName() As String) As Long

## Parameters

NumberNames

The number of area stiffness modifier names retrieved by the program.

MyName

This is a one-dimensional array of area stiffness modifier names. The MyName array is created as a dynamic, zero-based array by the API user:



```
Dim MyName() as String
```



The array is dimensioned to (NumberNames - 1) inside the SAP2000 program, filled with the names, and returned to the API user.

## Remarks

This function retrieves the names of all defined area stiffness modifiers.

The function returns zero if the names are successfully retrieved; otherwise it returns nonzero values.

## VBA Example



```
Sub GetAreaStiffnessModifierNames()

   'dimension variables

      Dim SapObject as cOAPI

      Dim SapModel As cSapModel

      Dim ret As Long

      Dim i As Long

      Dim Value() As Double

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



   'add new stiffness modifier

      ReDim Value(9)

      For i = 0 To 9

         Value(i) = 1

      Next i

      Value(5) = 2

      ret = SapModel.NamedAssign.ModifierArea.SetModifiers("AMOD1", Value)

      Value(5) = 1.5

      ret = SapModel.NamedAssign.ModifierArea.SetModifiers("AMOD2", Value)



   'get area stiffness modifier names

      ret = SapModel.NamedAssign.ModifierArea.GetNameList(NumberNames, MyName)



   'close Sap2000

      SapObject.ApplicationExit False

      Set SapModel = Nothing

      Set SapObject = Nothing

End Sub
```



## Release Notes

Initial release in version 12.00.

## See Also

# GetNameList

## Syntax

SapObject.SapModel.NamedAssign.ModifierCable.GetNameList

## VB6 Procedure

Function GetNameList(ByRef NumberNames As Long, ByRef MyName() As String) As Long

## Parameters

NumberNames

The number of cable property modifier names retrieved by the program.

MyName

This is a one-dimensional array of cable property modifier names. The MyName array is created as a dynamic, zero-based array by the API user:



```
Dim MyName() as String
```



The array is dimensioned to (NumberNames - 1) inside the SAP2000 program, filled with the names, and returned to the API user.

## Remarks

This function retrieves the names of all defined cable property modifiers.

The function returns zero if the names are successfully retrieved; otherwise it returns nonzero.

## VBA Example



```
Sub GetCablePropModifierNames()

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

      ret = SapModel.File.New2DFrame(PortalFrame, 2, 144, 2, 288)



   'add new property modifier

      ReDim Value(2)

      For i = 0 To 2

         Value(i) = 1

      Next i

      Value(1) = 2

      ret = SapModel.NamedAssign.ModifierCable.SetModifiers("CMOD1", Value)

      Value(1) = 1.5

      ret = SapModel.NamedAssign.ModifierCable.SetModifiers("CMOD2", Value)



   'get cable property modifier names

      ret = SapModel.NamedAssign.ModifierCable.GetNameList(NumberNames, MyName)



   'close Sap2000

      SapObject.ApplicationExit False

      Set SapModel = Nothing

      Set SapObject = Nothing

End Sub
```



## Release Notes

Initial release in version 12.00.

## See Also

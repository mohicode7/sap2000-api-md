# GetModifiers

## Syntax

SapObject.SapModel.NamedAssign.ModifierArea.GetModifiers

## VB6 Procedure

Function GetModifiers(ByVal Name As String, ByRef Value() As Double) As Long

## Parameters

Name

The name of an existing area stiffness modifier.

Value

This is an array of ten unitless modifiers.

Value(0) = Membrane f11 modifier

Value(1) = Membrane f22 modifier

Value(2) = Membrane f12 modifier

Value(3) = Bending m11 modifier

Value(4) = Bending m22 modifier

Value(5) = Bending m12 modifier

Value(6) = Shear v13 modifier

Value(7) = Shear v23 modifier

Value(8) = Mass modifier

Value(9) = Weight modifier

## Remarks

This function retrieves the modifier assignments for an area stiffness modifier. The default value for all modifier values is one.

The function returns zero if the modifier assignments are successfully retrieved; otherwise it returns a nonzero value.

## VBA Example



```
Sub GetAreaPropModifierValues()

   'dimension variables

      Dim SapObject as cOAPI

      Dim SapModel As cSapModel

      Dim ret As Long

      Dim i As Long

      Dim Value() As Double



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



   'define modifiers

      ReDim Value(9)

      For i = 0 To 9

         Value(i) = 1

      Next i

      Value(5) = 2

      ret = SapModel.NamedAssign.ModifierArea.SetModifiers("AMOD1", Value)



   'get modifier values

      ReDim Value(9)

      ret = SapModel.NamedAssign.ModifierArea.GetModifiers("AMOD1", Value)



   'close Sap2000

      SapObject.ApplicationExit False

      Set SapModel = Nothing

      Set SapObject = Nothing

End Sub
```



## Release Notes

Initial release in version 12.00.

## See Also

[SetModifiers](SetModifiers_%7BArea_Modifiers%7D.md)

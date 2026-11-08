# GetModifiers

## Syntax

SapObject.SapModel.AreaObj.GetModifiers

## VB6 Procedure

Function GetModifiers(ByVal Name As String, ByRef Value() As Double) As Long

## Parameters

Name

The name of an existing area object.

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

This function retrieves the modifier assignment for area objects. The default value for all modifiers is one.

The function returns zero if the modifier assignments are successfully retrieved; otherwise it returns a nonzero value.

## VBA Example



```
Sub GetAreaObjectModifiers()

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



   'assign modifiers

      ReDim Value(9)

      For i = 0 To 9

         Value(i) = 1

      Next i

      Value(0) = 0.01

      ret = SapModel.AreaObj.SetModifiers("ALL", Value, Group)



   'get modifiers

      ReDim Value(9)

      ret = SapModel.AreaObj.GetModifiers("3", Value)



   'close Sap2000

      SapObject.ApplicationExit False

      Set SapModel = Nothing

      Set SapObject = Nothing

End Sub
```



## Release Notes

Initial release in version 11.00.

## See Also

[SetModifiers](SetModifiers_%7BArea_Object%7D.md)

[DeleteModifiers](DeleteModifiers_%7BArea_Object%7D.md)

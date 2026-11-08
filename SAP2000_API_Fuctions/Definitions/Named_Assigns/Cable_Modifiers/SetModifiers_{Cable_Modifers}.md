# SetModifiers

## Syntax

SapObject.SapModel.NamedAssign.ModifierCable.SetModifiers

## VB6 Procedure

Function SetModifiers(ByVal Name As String, ByRef Value() As Double) As Long

## Parameters

Name

The name of a new or existing cable property modifier.

Value

This is an array of three unitless modifiers.

Value(0) = Cross sectional area modifier

Value(1) = Mass modifier

Value(2) = Weight modifier

## Remarks

This function defines a named cable property modifier. The default value for all modifier values is one.

The function returns zero if the modifier is successfully defined; otherwise it returns a nonzero value.

## VBA Example



```
Sub DefineCablePropModifiers()

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

      ret = SapModel.File.New2DFrame(PortalFrame, 3, 124, 3, 200)



   'define modifiers

      ReDim Value(2)

      For i = 0 To 2

         Value(i) = 1

      Next i

      Value(1) = 2

      ret = SapModel.NamedAssign.ModifierCable.SetModifiers("CMOD1", Value)



   'close Sap2000

      SapObject.ApplicationExit False

      Set SapModel = Nothing

      Set SapObject = Nothing

End Sub
```



## Release Notes

Initial release in version 12.00.

## See Also

[GetModifiers](GetModifiers_%7BCable_Modifers%7D.md)

# SetModifiers

## Syntax

SapObject.SapModel.NamedAssign.ModifierFrame.SetModifiers

## VB6 Procedure

Function SetModifiers(ByVal Name As String, ByRef Value() As Double) As Long

## Parameters

Name

The name of a new or existing frame property modifier.

Value

This is an array of eight unitless modifiers.

Value(0) = Cross sectional area modifier

Value(1) = Shear area in local 2 direction modifier

Value(2) = Shear area in local 3 direction modifier

Value(3) = Torsional constant modifier

Value(4) = Moment of inertia about local 2 axis modifier

Value(5) = Moment of inertia about local 3 axis modifier

Value(6) = Mass modifier

Value(7) = Weight modifier

## Remarks

This function defines a named frame property modifier. The default value for all modifier values is one.

The function returns zero if the modifier is successfully defined; otherwise it returns a nonzero value.

## VBA Example



```
Sub DefineFramePropModifiers()

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

      ReDim Value(7)

      For i = 0 To 7

         Value(i) = 1

      Next i

      Value(5) = 100

      ret = SapModel.NamedAssign.ModifierFrame.SetModifiers("FMOD1", Value)



   'close Sap2000

      SapObject.ApplicationExit False

      Set SapModel = Nothing

      Set SapObject = Nothing
End Sub
```



## Release Notes

Initial release in version 12.00.

## See Also

[GetModifiers](GetModifiers_%7BFrame_Modifiers%7D.md)

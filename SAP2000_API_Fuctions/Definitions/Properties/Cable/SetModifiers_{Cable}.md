# SetModifiers

## Syntax

SapObject.SapModel.PropCable.SetModifiers

## VB6 Procedure

Function SetModifiers(ByVal Name As String, ByRef Value() As Double) As Long

## Parameters

Name

The name of an existing cable property.

Value

This is an array of 3 unitless modifiers.

Value(0) = Cross sectional area modifier

Value(1) = Mass modifier

Value(2) = Weight modifier

## Remarks

This function assigns property modifiers to a cable property. The default value for all modifiers is one.

The function returns zero if the modifiers are successfully assigned; otherwise it returns a nonzero value.

## VBA Example



```
Sub AssignCablePropModifiers()

   'dimension variables

      Dim SapObject as cOAPI

      Dim SapModel As cSapModel

      Dim ret As Long

      Dim Name As String

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

      ret = SapModel.File.New2DFrame(PortalFrame, 2, 144, 2, 288)



   'add material

      ret = SapModel.PropMaterial.AddQuick(Name, MATERIAL_TENDON, , , , , , MATERIAL_TENDON_SUBTYPE_ASTM_A416Gr270)



   'set new cable property

      ret = SapModel.PropCable.SetProp("C1", Name, 2.25)



   'assign modifiers

      ReDim Value(2)

      For i = 0 To 2

         Value(i) = 1

      Next i

      Value(1) = 2

      Value(2) = 2

      ret = SapModel.PropCable.SetModifiers("C1", Value)



   'close Sap2000

      SapObject.ApplicationExit False

      Set SapModel = Nothing

      Set SapObject = Nothing

End Sub
```



## Release Notes

Initial release in version 11.02.

## See Also

[GetModifiers](GetModifiers_%7BCable%7D.md)

# GetModifiers

## Syntax

SapObject.SapModel.CableObj.GetModifiers

## VB6 Procedure

Function GetModifiers(ByVal Name As String, ByRef Value() As Double) As Long

## Parameters

Name

The name of an existing cable object.

Value

This is an array of three unitless modifiers.

Value(0) = Cross sectional area modifier

Value(1) = Mass modifier

Value(2) = Weight modifier

## Remarks

This function retrieves the cable modifier assignment for cable objects. The default value for all modifiers is one.

The function returns zero if the modifier assignments are successfully retrieved, otherwise it returns a nonzero value.

## VBA Example



```
Sub GetCableModifiers()

   'dimension variables

      Dim SapObject as cOAPI

      Dim SapModel As cSapModel

      Dim ret As Long

      Dim i As Long

      Dim Value() As Double

      Dim Name As String



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



   'add cable object by points

      ret = SapModel.CableObj.AddByPoint("1", "6", Name)



   'set cable data

      ret = SapModel.CableObj.SetCableData(Name, 7, 1, 0, 0, 24)



   'assign modifiers

      ReDim Value(2)

      For i = 0 To 2

         Value(i) = 1

      Next i

      Value(0) = 100

      ret = SapModel.CableObj.SetModifiers(Name, Value)



   'get modifiers

      ReDim Value(2)

      ret = SapModel.CableObj.GetModifiers(Name, Value)



   'close Sap2000

      SapObject.ApplicationExit False

      Set SapModel = Nothing

      Set SapObject = Nothing

End Sub
```



## Release Notes

Initial release in version 11.00.

## See Also

[SetModifiers](SetModifiers_%7BCable_Object%7D.md)

[DeleteModifiers](DeleteModifiers_%7BCable_Object%7D.md)

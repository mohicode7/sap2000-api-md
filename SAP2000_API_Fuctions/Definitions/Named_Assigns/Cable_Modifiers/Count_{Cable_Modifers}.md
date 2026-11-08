# Count

## Syntax

SapObject.SapModel.NamedAssign.ModifierCable.Count

## VB6 Procedure

Function Count() As Long

## Parameters

None

## Remarks

This function returns the total number of defined cable property modifiers in the model.

## VBA Example



```
Sub CountCablePropModifiers()

   'dimension variables

      Dim SapObject as cOAPI

      Dim SapModel As cSapModel

      Dim ret As Long

      Dim Count As Long

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



   'add new property modifier

      ReDim Value(2)

      For i = 0 To 2

         Value(i) = 1

      Next i

      Value(1) = 2

      ret = SapModel.NamedAssign.ModifierCable.SetModifiers("CMOD1", Value)

      Value(1) = 1.5

      ret = SapModel.NamedAssign.ModifierCable.SetModifiers("CMOD2", Value)



   'return number of defined property modifiers

      Count = SapModel.NamedAssign.ModifierCable.Count



   'close Sap2000

      SapObject.ApplicationExit False

      Set SapModel = Nothing

      Set SapObject = Nothing

End Sub
```



## Release Notes

Initial release in version 12.00.

## See Also

# Count

## Syntax

SapObject.SapModel.NamedAssign.ModifierArea.Count

## VB6 Procedure

Function Count() As Long

## Parameters

None

## Remarks

This function returns the total number of defined area stiffness modifiers in the model.

## VBA Example



```
Sub CountAreaStiffnessModifiers()

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



   'return number of defined stiffness modifiers

      Count = SapModel.NamedAssign.ModifierArea.Count



   'close Sap2000

      SapObject.ApplicationExit False

      Set SapModel = Nothing

      Set SapObject = Nothing

End Sub
```



## Release Notes

Initial release in version 12.00.

## See Also

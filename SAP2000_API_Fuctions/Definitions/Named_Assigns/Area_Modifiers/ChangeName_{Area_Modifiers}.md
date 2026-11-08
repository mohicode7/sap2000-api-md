# ChangeName

## Syntax

SapObject.SapModel.NamedAssign.ModifierArea.ChangeName

## VB6 Procedure

Function ChangeName(ByVal Name As String, ByVal NewName As String) As Long

## Parameters

Name

The existing name of a defined area stiffness modifier.

NewName

The new name for the area stiffness modifier.

## Remarks

This function changes the name of an existing area stiffness modifier.

The function returns zero if the new name is successfully applied; otherwise it returns a nonzero value.

## VBA Example



```
Sub ChangeAreaStiffnessModifierName()

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



   'add new stiffness modifier

      ReDim Value(9)

      For i = 0 To 9

         Value(i) = 1

      Next i

      Value(5) = 100

      ret = SapModel.NamedAssign.ModifierArea.SetModifiers("AMOD1", Value)



   'change name of stiffness modifier

      ret = SapModel.NamedAssign.ModifierArea.ChangeName("AMOD1", "MyModifier")



   'close Sap2000

      SapObject.ApplicationExit False

      Set SapModel = Nothing

      Set SapObject = Nothing

End Sub
```



## Release Notes

Initial release in version 12.00.

## See Also

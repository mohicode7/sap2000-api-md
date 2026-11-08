# Delete

## Syntax

SapObject.SapModel.NamedAssign.ModifierArea.Delete

## VB6 Procedure

Function Delete(ByVal Name As String) As Long

## Parameters

Name

The name of an existing area stiffness modifier.

## Remarks

The function deletes a specified area stiffness modifier.

The function returns zero if the modifier is successfully deleted; otherwise it returns a nonzero value. It returns an error if the specified modifier can not be deleted; for example, if it is currently used by a staged construction load case.

## VBA Example



```
Sub DeleteAreaStiffnessModifier()

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

      Value(5) = 2

      ret = SapModel.NamedAssign.ModifierArea.SetModifiers("AMOD1", Value)

      Value(5) = 1.5

      ret = SapModel.NamedAssign.ModifierArea.SetModifiers("AMOD2", Value)



   'delete stiffness modifier

      ret = SapModel.NamedAssign.ModifierArea.Delete("AMOD1")



   'close Sap2000

      SapObject.ApplicationExit False

      Set SapModel = Nothing

      Set SapObject = Nothing

End Sub
```



## Release Notes

Initial release in version 12.00.

## See Also

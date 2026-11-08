# ChangeName

## Syntax

SapObject.SapModel.NamedAssign.ModifierFrame.ChangeName

## VB6 Procedure

Function ChangeName(ByVal Name As String, ByVal NewName As String) As Long

## Parameters

Name

The existing name of a defined frame property modifier.

NewName

The new name for the frame property modifier.

## Remarks

This function changes the name of an existing frame property modifier.

The function returns zero if the new name is successfully applied; otherwise it returns a nonzero value.

## VBA Example



```
Sub ChangeFramePropModifierName()

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



   'add new property modifier

      ReDim Value(7)

      For i = 0 To 7

         Value(i) = 1

      Next i

      Value(5) = 100

      ret = SapModel.NamedAssign.ModifierFrame.SetModifiers("FMOD1", Value)



   'change name of property modifier

      ret = SapModel.NamedAssign.ModifierFrame.ChangeName("FMOD1", "MyModifier")



   'close Sap2000

      SapObject.ApplicationExit False

      Set SapModel = Nothing

      Set SapObject = Nothing

End Sub
```



## Release Notes

Initial release in version 12.00.

## See Also

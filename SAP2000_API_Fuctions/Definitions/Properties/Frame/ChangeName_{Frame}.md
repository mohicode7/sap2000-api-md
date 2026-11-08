# ChangeName

## Syntax

SapObject.SapModel.PropFrame.ChangeName

## VB6 Procedure

Function ChangeName(ByVal Name As String, ByVal NewName As String) As Long

## Parameters

Name

The existing name of a defined frame section property.

NewName

The new name for the frame section property.

## Remarks

This function changes the name of an existing frame section property.

The function returns zero if the new name is successfully applied; otherwise it returns a nonzero value.

## VBA Example



```
Sub ChangeFramePropName()

   'dimension variables

      Dim SapObject as cOAPI

      Dim SapModel As cSapModel

      Dim ret As Long



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



   'change name of frame section property

      ret = SapModel.PropFrame.ChangeName("FSEC1", "MySection")



   'close Sap2000

      SapObject.ApplicationExit False

      Set SapModel = Nothing

      Set SapObject = Nothing

End Sub
```



## Release Notes

Initial release in version 11.02.

## See Also

# NewBlank

## Syntax

SapObject.SapModel.File.NewBlank

## VB6 Procedure

Function NewBlank() As Long

## Parameters

None

## Remarks

Do not use this function to add to an existing model. This function should be used only for creating a new model and typically would be preceded by calls to ApplicationStart or InitializeNewModel.

The function returns zero if the new blank model is successfully created, otherwise it returns a nonzero value.

## VBA Example



```
Sub NewBlankModel()

   'dimension variables

      Dim SapObject as cOAPI

      Dim SapModel As cSapModel

      Dim ret as long



   'create Sap2000 object

      Set SapObject = CreateObject("CSI.SAP2000.API.SapObject")



   'start Sap2000 application

      SapObject.ApplicationStart



   'create SapModel object

      Set SapModel = SapObject.SapModel



   'initialize model

      ret = SapModel.InitializeNewModel



   'create a blank model from template

      ret = SapModel.File.NewBlank



   'close Sap2000

      SapObject.ApplicationExit False

      Set SapModel = Nothing

      Set SapObject = Nothing

End Sub
```



## Release Notes

Initial release in version 11.00.

## See Also

[NewBeam](NewBeam.md)

[New2DFrame](New2DFrame.md)

[New3DFrame](New3DFrame.md)

[NewSolidBlock](NewSolidBlock.md)

[NewWall](NewWall.md)

[ApplicationStart](../General_Functions/ApplicationStart.md)

[InitializeNewModel](../General_Functions/InitializeNewModel.md)

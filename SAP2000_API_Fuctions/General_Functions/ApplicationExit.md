# ApplicationExit

## Syntax

SapObject.ApplicationExit

## VB6 Procedure

Function ApplicationExit(ByVal FileSave As Boolean) As Long

## Parameters

FileSave

If this item is True the existing model file is saved prior to closing Sap2000.

## Remarks

If the model file is saved then it is saved with its current name. You should set the Sap2000 object variable to nothing after calling this function.

This function returns zero if the function succeeds and nonzero if it fails.

## VBA Example



```
Sub ExitExample()

   'dimension variables

      Dim SapObject as cOAPI

      Dim SapModel As cSapModel

      Dim ret as Long



   'create Sap2000 object

      Set SapObject = CreateObject("CSI.SAP2000.API.SapObject")



   'start Sap2000 application

      SapObject.ApplicationStart



   'create SapModel object

      Set SapModel = SapObject.SapModel



   'initialize model

      ret = SapModel.InitializeNewModel(kip_ft_F)



   'create a 3D frame model from template

      ret = SapModel.File.New3DFrame(BeamSlab, 3, 12, 3, 28, 2, 36)   



   'close Sap2000

      SapObject.ApplicationExit False

      Set SapModel = Nothing

      Set SapObject = Nothing

End Sub
```



## Release Notes

Initial release in version 11.00.

## See Also

[ApplicationStart](ApplicationStart.md)

# ClearSelection

## Syntax

SapObject.SapModel.SelectObj.ClearSelection

## VB6 Procedure

Function ClearSelection() As Long

## Parameters

None

## Remarks

This function deselects all objects in the model. It returns zero if the selection status is successfully set, otherwise it returns nonzero.

## VBA Example



```
Sub ClearSelection()

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



   'create model from template

      ret = SapModel.File.New2DFrame(ConcentricBraced, 3, 124, 3, 200)



   'select all

      ret = SapModel.SelectObj.All



   'clear selection

      ret = SapModel.SelectObj.ClearSelection



   'close Sap2000

      SapObject.ApplicationExit False

      Set SapModel = Nothing

      Set SapObject = Nothing

End Sub
```



## Release Notes

Initial release in version 11.00.

## See Also

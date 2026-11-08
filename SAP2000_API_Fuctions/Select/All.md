# All

## Syntax

SapObject.SapModel.SelectObj.All

## VB6 Procedure

Function All(Optional ByVal DeSelect As Boolean = False) As Long

## Parameters

DeSelect

The item is False if objects are to be selected and True if they are to be deselected.

## Remarks

This function selects or deselects all objects in the model.

This function returns zero if the selection is successfully completed, otherwise it returns nonzero.

## VBA Example



```
Sub SelectAll()

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



   'deselect all

      ret = SapModel.SelectObj.All(True)



   'close Sap2000

      SapObject.ApplicationExit False

      Set SapModel = Nothing

      Set SapObject = Nothing

End Sub
```



## Release Notes

Initial release in version 11.00.

## See Also

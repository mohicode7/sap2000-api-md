# SetOptionModeShape

## Syntax

SapObject.SapModel.Results.Setup.SetOptionModeShape

## VB6 Procedure

Function SetOptionModeShape(ByVal ModeShapeStart As Long, ByVal ModeShapeEnd As Long, Optional ByVal ModeShapesAll As Boolean = False) As Long

## Parameters

ModeShapeStart

The first mode for which results are reported when the ModeShapesAll item is False.

ModeShapeEnd

The last mode for which results are reported when the ModeShapesAll item is False.

ModeShapesAll

If this item is True, results are reported for all calculated modes. If it is False, results are reported for the modes indicated by the ModeShapeStart and ModeShapeEnd items.

## Remarks

This function sets the modes for which mode shape results are reported.

The function returns 0 if the modes are successfully set, otherwise it returns nonzero.

## VBA Example



```
Sub SetModeShapeData()

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



   'set mode shape data

      ret = SapModel.Results.Setup.SetOptionModeShape(1, 3)



   'close Sap2000

      SapObject.ApplicationExit False

      Set SapModel = Nothing

      Set SapObject = Nothing

End Sub
```



## Release Notes

Initial release in version 11.00.

## See Also

[GetOptionModeShape](GetOptionModeShape.md)

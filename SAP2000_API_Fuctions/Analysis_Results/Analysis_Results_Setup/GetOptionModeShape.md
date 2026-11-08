# GetOptionModeShape

## Syntax

SapObject.SapModel.Results.Setup.GetOptionModeShape

## VB6 Procedure

Function GetOptionModeShape(ByRef ModeShapeStart As Long, ByRef ModeShapeEnd As Long, ByRef ModeShapesAll As Boolean) As Long

## Parameters

ModeShapeStart

The first mode for which results are reported when the ModeShapesAll item is False.

ModeShapeEnd

The last mode for which results are reported when the ModeShapesAll item is False.

ModeShapesAll

If this item is True, results are reported for all calculated modes. If it is False, results are reported for the modes indicated by the ModeShapeStart and ModeShapeEnd items.

## Remarks

This function retrieves the modes for which mode shape results are reported.

The function returns 0 if the modes are successfully retrieved, otherwise it returns nonzero.

## VBA Example



```
Sub GetModeShapeData()

   'dimension variables

      Dim SapObject as cOAPI

      Dim SapModel As cSapModel

      Dim ret As Long

      Dim ModeShapeStart As Long

      Dim ModeShapeEnd As Long

      Dim ModeShapesAll As Boolean



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



   'get buckling mode data

      ret = SapModel.Results.Setup.GetOptionModeShape(ModeShapeStart, ModeShapeEnd, ModeShapesAll)



   'close Sap2000

      SapObject.ApplicationExit False

      Set SapModel = Nothing

      Set SapObject = Nothing

End Sub
```



## Release Notes

Initial release in version 11.00.

## See Also

[SetOptionModeShape](SetOptionModeShape.md)

# Count

## Syntax

Sap2000.LineElm.Count

## VB6 Procedure

Function Count() As Long

## Parameters

None

## Remarks

This function returns the total number of line elements in the analysis model.

## VBA Example



```
Sub CountLineElements()

   'dimension variables

      Dim SapObject as cOAPI

      Dim SapModel As cSapModel

      Dim ret As Long

      Dim Count As Long



   'create Sap2000 object

      Set SapObject = CreateObject("CSI.SAP2000.API.SapObject")



   'start Sap2000 application

      SapObject.ApplicationStart



   'create SapModel object

      Set SapModel = SapObject.SapModel



   'initialize model

      ret = SapModel.InitializeNewModel



   'create model from template

      ret = SapModel.File.New2DFrame(PortalFrame, 2, 144, 2, 288)



   'assign automesh options

      ret = SapModel.FrameObj.SetAutoMesh("ALL", True, True, True, 2, 0, Group)



   'create analysis model

      ret = SapModel.Analyze.CreateAnalysisModel



   'return number of line elements

      Count = SapModel.LineElm.Count



   'close Sap2000

      SapObject.ApplicationExit False

      Set SapModel = Nothing

      Set SapObject = Nothing

End Sub
```



## Release Notes

Initial release in version 11.00.

## See Also

# RunAnalysis

## Syntax

SapObject.SapModel.Analyze.RunAnalysis

## VB6 Procedure

Function RunAnalysis() As Long

## Parameters

None

## Remarks

This function runs the analysis. The analysis model is automatically created as part of this function.

The function returns zero if the analysis model is successfully run, otherwise it returns a nonzero value.

IMPORTANT NOTE: Your model must have a file path defined before running the analysis. If the model is opened from an existing file, a file path is defined. If the model is created from scratch, the File.Save function must be called with a file name before running the analysis. Saving the file creates the file path.

## VBA Example



```
Sub RunSapAnalysisModel()

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



   'save model

      ret = SapModel.File.Save("C:\SapAPI\x.sdb")



   'run model (this will create the analysis model)

      ret = SapModel.Analyze.RunAnalysis



   'close Sap2000

      SapObject.ApplicationExit False

      Set SapModel = Nothing

      Set SapObject = Nothing

End Sub
```



## Release Notes

Initial release in version 11.00.

## See Also

[File.Save](../File/Save.md)

[BridgeObj.SetBridgeUpdateForAnalysisFlag](../Definitions/Bridge_Objects/SetBridgeUpdateForAnalysisFlag.md)

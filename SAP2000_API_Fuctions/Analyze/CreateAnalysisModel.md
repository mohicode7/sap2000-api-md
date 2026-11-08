# CreateAnalysisModel

## Syntax

SapObject.SapModel.Analyze.CreateAnalysisModel

## VB6 Procedure

Function CreateAnalysisModel() As Long

## Parameters

None

## Remarks

This function creates the analysis model. If the analysis model is already created and current, nothing is done.

The function returns zero if the analysis model is successfully created or it already exists and is current, otherwise it returns a nonzero value.

It is not necessary to call this function before running an analysis. The analysis model is automatically created, if necessary, when the model is run.

## VBA Example



```
Sub CreateSapAnalysisModel()

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



   'create the analysis model

      ret = SapModel.Analyze.CreateAnalysisModel



   'close Sap2000

      SapObject.ApplicationExit False

      Set SapModel = Nothing

      Set SapObject = Nothing

End Sub
```



## Release Notes

Initial release in version 11.00.

## See Also

[RunAnalysis](RunAnalysis.md)

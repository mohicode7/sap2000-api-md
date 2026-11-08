# MergeAnalysisResults

## Syntax

SapObject.SapModel.Analyze.MergeAnalysisResults

## VB6 Procedure

Function MergeAnalysisResults(ByVal FileName As String)
As Long

## Parameters

FileName

The full path of a model file from which the analysis
results are to be merged.

## Remarks

See “Merging Analysis Results” section in program help
file for requirements and limitations.

The analysis model is automatically created as part
of this function.

The function returns zero if analysis results are successfully
merged, otherwise it returns a nonzero value.

IMPORTANT NOTE: Your model must have a file path defined
before merging analysis results. If the model is opened from an existing
file, a file path is defined. If the model is created from scratch, the
File.Save function must be called with a file name before merging analysis
results. Saving the file creates the file path.

## VBA Example



```
Sub MergeSapAnalysisResults()

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

ret = SapModel.File.Save("C:\SapAPI\source\model.sdb")



'run model (this will create the analysis model)

ret = SapModel.Analyze.RunAnalysis



'initialize a new model

ret = SapModel.InitializeNewModel



'create the same model from template

ret = SapModel.File.New2DFrame(PortalFrame, 3, 124, 3, 200)



'merge analysis results

ret = SapModel.Analyze.SetSolverOption_2(1, 2, 0)
ret = SapModel.File.Save("C:\SapAPI\target\model.sdb")

ret = SapModel.Analyze.MergeAnalysisResults("C:\SapAPI\source\model.sdb")



'close Sap2000

SapObject.ApplicationExit False

Set SapModel = Nothing

Set SapObject = Nothing 
End Sub
```



## Release Notes

Initial release in version 22.1.0.

## See Also

[File.Save](../File/Save.md)

[Analyze.RunAnalysis](RunAnalysis.md)

[Analyze.SetSolverOption\_2](../Obsolete_Functions/NEW_-_Analyze.SetSolverOption_2.md)

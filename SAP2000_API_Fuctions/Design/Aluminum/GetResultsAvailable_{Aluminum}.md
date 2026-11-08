# GetResultsAvailable {Aluminum}

## Syntax

SapObject.SapModel.DesignAluminum.GetResultsAvailable

## VB6 Procedure

Function GetResultsAvailable() As Boolean

## Parameters

None

## Remarks

The function returns True if the aluminum frame design results are available, otherwise False.

## VBA Example



```
Sub GetResultsAvailable()
  'dimension variables
    Dim SapObject as cOAPI
    Dim SapModel As cSapModel
    Dim ret As Long
    Dim ResultsAvailable As Boolean
 
  'create Sap2000 object
    Set SapObject = CreateObject("CSI.SAP2000.API.SapObject")
 
  'start Sap2000 application
    SapObject.ApplicationStart
 
  'create SapModel object
    Set SapModel = SapObject.SapModel
 
  'initialize model
    ret = SapModel.InitializeNewModel
 
  'add aluminum formed material
    ret = SapModel.PropMaterial.AddQuick(Name, MATERIAL_ALUMINUM, , , MATERIAL_ALUMINUM_SUBTYPE_6061_T6)  
 
  'create new aluminum frame section property
    ret = SapModel.PropFrame.SetISection("AI", Name , 18, 6, 0.5, 0.3, 6, 0.5 )
 
  'create model from template
    ret = SapModel.File.New2DFrame(PortalFrame, 2, 144, 2, 288, True, "AI", "AI")
 
  'run analysis
    ret = SapModel.File.Save("C:\SapAPI\x.sdb")
    ret = SapModel.Analyze.RunAnalysis
 
  'start aluminum design
    ret = SapModel.DesignAluminum.StartDesign
 
  'check if design results are available
    ResultsAvailable = SapModel.DesignAluminum.GetResultsAvailable
 
  'close Sap2000
    SapObject.ApplicationExit.False
    Set SapModel = Nothing
    Set SapObject = Nothing
End Sub  
```



## Release Notes

Initial release in version 18.2.0.

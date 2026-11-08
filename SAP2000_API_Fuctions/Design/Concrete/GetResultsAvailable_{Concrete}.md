# GetResultsAvailable {Concrete}

## Syntax

SapObject.SapModel.DesignConcrete.GetResultsAvailable

## VB6 Procedure

Function GetResultsAvailable() As Boolean

## Parameters

None

## Remarks

The function returns True if the concrete frame design results are available, otherwise False.

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
 
  'create new concrete frame section property
    ret = SapModel.PropFrame.SetRectangle("R1", "4000Psi", 20, 12)
 
  'create model from template
    ret = SapModel.File.New2DFrame(PortalFrame, 2, 144, 2, 288, True, "R1", "R1")
 
  'run analysis
    ret = SapModel.File.Save("C:\SapAPI\x.sdb")
    ret = SapModel.Analyze.RunAnalysis
 
  'start concrete design
    ret = SapModel.DesignConcrete.StartDesign
 
  'check if design results are available
    ResultsAvailable = SapModel.DesignConcrete.GetResultsAvailable
 
  'close Sap2000
    SapObject.ApplicationExit.False
    Set SapModel = Nothing
    Set SapObject = Nothing
End Sub  
```



## Release Notes

Initial release in version 18.2.0.

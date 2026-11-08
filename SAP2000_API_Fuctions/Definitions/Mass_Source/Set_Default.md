# Set Default

## Syntax

SapObject.SapModel.SourceMass.SetDefault

## VB6 Procedure

Function SetDefault(ByVal Name As String) As Long

## Parameters

Name

The name of the mass source to be flagged as the default mass source.

## **Remarks**

This function sets the default mass source.

The function returns zero if the mass source is successfully flagged as default, otherwise it returns nonzero.  Only one mass source can be the default mass source so when this assignment is set all other mass sources are automatically set to have their IsDefault flag False.

## VBA Example



```
Sub SetDefaultMassSource()
'dimension variables
Dim SapObject as cOAPI
Dim SapModel As cSapModel
Dim ret As Long
Dim LoadPat(0) As String
Dim SF(0) As Double
 
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
 
'add a new mass source and do NOT make it the default mass source
LoadPat(0) = "DEAD"
SF(0) = 1.25
ret = SapModel.SourceMass.SetMassSource("MyMassSource", True, True, True, False, 1, LoadPat, SF)
 
'set MyMassSource as the default mass source
ret = SapModel.SourceMass.SetDefault("MyMassSource")
 
'close Sap2000
SapObject.ApplicationExit False
Set SapModel = Nothing
Set SapObject = Nothing
End Sub
```



## Release Notes

Initial release in version 16.0.0.

## See Also

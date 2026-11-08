# Get Default {Mass Source}

## Syntax

SapObject.SapModel.SourceMass.GetDefault

## VB6 Procedure

Function GetDefault(ByRef Name As String) As Long

## Parameters

Name

The name of the default mass source.

## **Remarks**

This function retrieves the default mass source.

The function returns zero if the default mass source is successfully retrieved, otherwise it returns nonzero.

## VBA Example



```
Sub GetDefaultMassSource()
'dimension variables
Dim SapObject as cOAPI
Dim SapModel As cSapModel
Dim ret As Long
Dim LoadPat(0) As String
Dim SF(0) As Double
Dim DefaultMassSource As String
 
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
 
'add a new mass source and make it the default mass source
LoadPat(0) = "DEAD"
SF(0) = 1.25
ret = SapModel.SourceMass.SetMassSource("MyMassSource", True, True, True, True, 1, LoadPat, SF)
 
'get the default mass source
ret = SapModel.SourceMass.GetDefault(DefaultMassSource)
 
'close Sap2000
SapObject.ApplicationExit False
Set SapModel = Nothing
Set SapObject = Nothing
End Sub
```



## Release Notes

Initial release in version 16.0.0.

## See Also

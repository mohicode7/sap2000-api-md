# Set Mass Source

## Syntax

SapObject.SapModel.SourceMass.SetMassSource

## VB6 Procedure

Function SetMassSource(ByVal Name As String, ByVal MassFromElements As Boolean, ByVal MassFromMasses As Boolean, ByVal MassFromLoads As Boolean, ByVal IsDefault As Boolean, ByVal NumberLoads As Long, ByRef LoadPat() As String, ByRef SF() As Double) As Long

## Parameters

Name

The mass source name. If a mass source with this name already exists then the mass source is reinitialized with the new data.  All previous data assigned to the mass source is lost. If a mass source with this name does not exist then a new mass source is added.

MassFromElements

If this item is True then element self mass is included in the mass.

MassFromMasses

If this item is True then assigned masses are included in the mass.

MassFromLoads

If this item is True then specified load patterns are included in the mass.

IsDefault

If this item is True then the mass source is the default mass source.  Only one mass source can be the default mass source so when this assignment is True all other mass sources are automatically set to have the IsDefault flag False.

NumberLoads

The number of load patterns specified for the mass source.  This item is only applicable when the MassFromLoads item is True.

LoadPat

This is an array of load pattern names specified for the mass source.

SF

This is an array of load pattern multipliers specified for the mass source.

## **Remarks**

This function adds a new mass source to the model or reinitializes an existing mass source.

The function returns zero if the mass source is successfully added or initialized, otherwise it returns a nonzero value.

## VBA Example



```
Sub SetMassSource()
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
 
'add a new mass source and make it the default mass source
LoadPat(0) = "DEAD"
SF(0) = 1.25
ret = SapModel.SourceMass.SetMassSource("MyMassSource", True, True, True, True, 1, LoadPat, SF)
 
'close Sap2000
SapObject.ApplicationExit False
Set SapModel = Nothing
Set SapObject = Nothing
End Sub
```



## Release Notes

Initial release in version 16.0.0.

## See Also

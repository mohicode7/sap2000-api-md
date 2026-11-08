# Get Mass Source

## Syntax

SapObject.SapModel.SourceMass.GetMassSource

## VB6 Procedure

Function GetMassSource (ByVal Name As String, ByRef MassFromElements As Boolean, ByRef MassFromMasses As Boolean, ByRef MassFromLoads As Boolean, ByRef IsDefault As Boolean, ByRef NumberLoads As Long, ByRef LoadPat() As String, ByRef SF() As Double) As Long

## Parameters

Name

The mass source name.

MassFromElements

If this item is True then element self mass is included in the mass.

MassFromMasses

If this item is True then assigned masses are included in the mass.

MassFromLoads

If this item is True then specified load patterns are included in the mass.

IsDefault

If this item is True then the mass source is the default mass source.

NumberLoads

The number of load patterns specified for the mass source.  This item is only applicable when the MassFromLoads item is True.

LoadPat

This is an array of load pattern names specified for the mass source.

SF

This is an array of load pattern multipliers specified for the mass source.

## **Remarks**

This function gets the mass source data for an existing mass source.

The function returns zero if the mass source data is successfully retrieved, otherwise it returns a nonzero value.

## VBA Example



```
Sub GetMassSource()
'dimension variables
Dim SapObject as cOAPI
Dim SapModel As cSapModel
Dim ret As Long
Dim MassFromElements As Boolean
Dim MassFromMasses As Boolean
Dim MassFromLoads As Boolean
Dim IsDefault As Boolean
Dim NumberLoads As Long
Dim LoadPat() As String
Dim SF() As Double
 
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
 
'get data for mass source MSSSRC1
ret = SapModel.SourceMass.GetMassSource("MSSSRC1", MassFromElements, MassFromMasses, MassFromLoads, IsDefault, NumberLoads, LoadPat, SF)
 
'close Sap2000
SapObject.ApplicationExit False
Set SapModel = Nothing
Set SapObject = Nothing
End Sub
```



## Release Notes

Initial release in version 16.0.0.

## See Also

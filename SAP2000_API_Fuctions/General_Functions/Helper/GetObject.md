# GetObject

## Syntax

Helper.GetObject

## VB6 Procedure

Function GetObject(ByVal typeName As String) As cOAPI

## Parameters

typeName

The program ID of the API object. Use “CSI.SAP2000.API.SapObject”
for SAP2000 and “CSI.CSiBridge.API.SapObject” for CSiBridge.

## Remarks

Attaches to the active running instance of SAP2000 and
returns an instance of SapObject if successful, nothing otherwise.

## VBA Example



```
Public Sub Example()
Dim myHelper As cHelper
Dim SapModel As cSapModel
Dim SapObject As cOAPI
Dim ret As Integer
 
‘create SapObject
Set myHelper = New Helper
Set SapObject = myHelper.GetObject("CSI.SAP2000.API.SapObject")
 
'create SapModel object
Set SapModel = SapObject.SapModel
 
'initialize model
ret = SapModel.InitializeNewModel()
 
'close SAP2000
ret = SapObject.ApplicationExit(False)
 
'clean up variables
Set SapModel = Nothing
Set SapObject = Nothing
Set myHelper = Nothing
End Sub
```



## Release Notes

Initial release in version 17.00

## See Also

[SapObject.AplicationExit](../ApplicationExit.md)

[SetAsActiveObject](../SetAsActiveObject.md)

[UnsetAsActiveObject](../UnsetAsActiveObject.md)

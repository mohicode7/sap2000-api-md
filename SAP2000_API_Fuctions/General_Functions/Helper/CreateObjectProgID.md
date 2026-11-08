# CreateObjectProgID

## Syntax

Helper.CreateObjectProgID

## VB6 Procedure

Function CreateObjectProgID(ByVal progID As String)
As cOAPI

## Parameters

progID

The program ID of the API object. Use “CSI.SAP2000.API.SapObject”
for SAP2000 and “CSI.CSiBridge.API.SapObject” for CSiBridge.

## Remarks

Starts SAP2000 and returns an instance of SapObject
if successful, nothing otherwise.

The function searches the registry for the newest installed
version of SAP2000 unless overridden with an environment variable containing
the full path to SAP2000.exe. The overriding environment variable is "CSI\_SAP2000\_API\_SapObject\_PATH"
for SAP2000 and "CSI\_CSiBridge\_API\_SapObject\_PATH" for CSiBridge.

## VBA Example



```
Public Sub Example()
Dim myHelper As cHelper
Dim SapModel As cSapModel
Dim SapObject As cOAPI
Dim ret As Integer
 
‘create SapObject
Set myHelper = New Helper
Set SapObject = myHelper.CreateObjectProgID("CSI.SAP2000.API.SapObject")
 
‘start SAP2000 application
ret = SapObject.ApplicationStart()
 
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

Initial release in version 19.10

## See Also

[SapObject.ApplicationStart](../ApplicationStart.md)

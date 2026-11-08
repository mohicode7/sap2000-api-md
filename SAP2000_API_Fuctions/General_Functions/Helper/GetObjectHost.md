# GetObjectHost

## Syntax

Helper.GetObjectHost

## VB6 Procedure

Function GetObjectHost(ByVal hostName As String, ByVal
progID As String) As cOAPI

## Parameters

hostName

The host computer name.

progID

The program ID of the API object. Use “CSI.SAP2000.API.SapObject”
for SAP2000 and “CSI.CSiBridge.API.SapObject” for CSiBridge.

## Remarks

Attaches to the active running instance of SAP2000 on
a given computer using the default TCP port and returns an instance of
SapObject if successful, nothing otherwise.

CSiAPIService.exe must be running on the host computer
and listening to the default TCP port for the call to succeed.

## VBA Example



```
Public Sub Example()
Dim myHelper As cHelper
Dim SapModel As cSapModel
Dim SapObject As cOAPI
Dim ret As Integer
 
‘create SapObject
Set myHelper = New Helper
Set SapObject = myHelper.GetObjectHost("RemoteServer", 
 "CSI.SAP2000.API.SapObject")
 
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

Initial release in version 22.10

## See Also

[SapObject.AplicationExit](../ApplicationExit.md)

[RemoteAPI](../../../Remote_API.md)

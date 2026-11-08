# GetObjectHostPort

## Syntax

Helper.GetObjectHostPort

## VB6 Procedure

Function GetObjectHostPort(ByVal hostName As String,
ByVal portNumber As Integer, ByVal progID As String) As cOAPI

## Parameters

hostName

The host computer name.

portNumber

The TCP port to connect to the host computer.

progID

The program ID of the API object. Use “CSI.SAP2000.API.SapObject”
for SAP2000 and “CSI.CSiBridge.API.SapObject” for CSiBridge.

## Remarks

Attaches to the active running instance of SAP2000 on
a given computer using the given TCP port. Returns an instance of SapObject
if successful, nothing otherwise.

CSiAPIService.exe must be running on the host computer
and listening to the given TCP port for the call to succeed.

## VBA Example



```
Public Sub Example()
Dim myHelper As cHelper
Dim SapModel As cSapModel
Dim SapObject As cOAPI
Dim ret As Integer
 
‘create SapObject
Set myHelper = New Helper
Set SapObject = myHelper.GetObjectHostPort("RemoteServer", 
 8500, "CSI.SAP2000.API.SapObject")
 
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

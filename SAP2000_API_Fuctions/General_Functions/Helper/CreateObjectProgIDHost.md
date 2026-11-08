# CreateObjectProgIDHost

## Syntax

Helper.CreateObjectProgIDHost

## VB6 Procedure

Function CreateObjectProgIDHost(ByVal hostName As String,
ByVal progID As String) As cOAPI

## Parameters

hostname

The host computer name.

progID

The program ID of the API object. Use “CSI.SAP2000.API.SapObject”
for SAP2000 and “CSI.CSiBridge.API.SapObject” for CSiBridge.

## Remarks

Starts SAP2000 on a given computer using the default
TCP port and returns an instance of SapObject if successful, nothing otherwise.

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
Set SapObject = myHelper.CreateObjectProgIDHost("RemoteServer", 
 "CSI.SAP2000.API.SapObject")
 
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

Initial release in version 22.10

## See Also

[SapObject.ApplicationStart](../ApplicationStart.md)

[RemoteAPI](../../../Remote_API.md)

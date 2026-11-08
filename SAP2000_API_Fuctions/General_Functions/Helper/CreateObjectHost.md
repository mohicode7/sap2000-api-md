# CreateObjectHost

## Syntax

Helper.CreateObjectHost

## VB6 Procedure

Function CreateObjectHost(ByVal hostName As String,
ByVal fullPath As String) As cOAPI

## Parameters

hostname

The host computer name.

fullPath

The full path to SAP2000.exe.

## Remarks

Starts SAP2000 at the given path on a given computer
using the default TCP port and returns an instance of SapObject if successful,
nothing otherwise.

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
Set SapObject = myHelper.CreateObjectHost("RemoteServer", 
 "C:\Program Files\Computers and Structures\SAP2000 22\SAP2000.exe")
 
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

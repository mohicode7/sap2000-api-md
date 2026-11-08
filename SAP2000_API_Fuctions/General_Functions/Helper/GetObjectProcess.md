# GetObjectProcess

## Syntax

Helper.GetObjectProcess

## VB6 Procedure

Function GetObjectProcess(ByVal typeName As String,
ByVal pid As Integer) As cOAPI

## Parameters

typeName

The program ID of the API object. Use “CSI.SAP2000.API.SapObject”
for SAP2000 and “CSI.CSiBridge.API.SapObject” for CSiBridge.

pid

The system-generated unique identifier (process ID)
of the desired instance of the program as shown in Windows Task Manager
and/or under “Tools” menu.

## Remarks

Attaches to the running instance of the program with
the given process ID and returns an instance of SapObject if successful,
nothing otherwise.

## VBA Example



```
Public Sub Example()
Dim myHelper As cHelper
Dim SapModel As cSapModel
Dim SapObject As cOAPI
Dim pid As Integer
Dim ret As Integer
 
‘create SapObject
Set myHelper = New Helper
pid = 15476 'set to process ID of the instance to attach 
 to
 
Set SapObject = myHelper.GetObjectProcess("CSI.SAP2000.API.SapObject", 
 pid)
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

Initial release in version 24.1.0

## See Also

[GetObject](GetObject.md)

[SetAsActiveObject](../SetAsActiveObject.md)

[UnsetAsActiveObject](../UnsetAsActiveObject.md)

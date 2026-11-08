# CreateObject

## Syntax

Helper.CreateObject

## VB6 Procedure

Function CreateObject(ByVal fullPath As String) As cOAPI

## Parameters

fullPath

The full path to SAP2000.exe.

## Remarks

Starts SAP2000 at the given path and returns an instance
of SapObject if successful, nothing otherwise.

## VBA Example



```
Public Sub Example()
Dim myHelper As cHelper
Dim SapModel As cSapModel
Dim SapObject As cOAPI
Dim ret As Long
 
'create SapObject
Set myHelper = New Helper
Set SapObject = myHelper.CreateObject("C:\Program 
 Files\Computers and Structures\SAP2000 22\SAP2000.exe")
 
'start SAP2000 application
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

Initial release in version 17.00

## See Also

[SapObject.ApplicationStart](../ApplicationStart.md)

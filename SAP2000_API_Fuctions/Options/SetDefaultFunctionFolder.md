# SetDefaultFunctionFolder

## Syntax

SapObject.SapModel.Options.SetDefaultFunctionFolder

## VB6 Procedure

Function SetDefaultFunctionFolder(ByRef Path As String)
As Long

## Parameters

Path

Full path to the default function folder.

## Remarks

This function sets the default folder to (recursively)
search if a function data file is not found in the specified path or in the
model folder. This is a global setting for the user and applies to all models.

The function returns zero if the data is successfully
set; otherwise it returns a nonzero value.

## VBA Example



```
Sub Example()
    Dim myHelper As cHelper
    Dim SapModel As cSapModel
    Dim SapObject As cOAPI
    Dim ret As Long
    Dim path As String
 
    'create SapObject
    Set myHelper = New Helper
    Set SapObject = myHelper.CreateObjectProgID("CSI.SAP2000.API.SapObject")
   
    'start SAP2000 application
    ret = SapObject.ApplicationStart()
   
    'create SapModel object
    Set SapModel = SapObject.SapModel
   
    'initialize model
    ret = SapModel.InitializeNewModel()
   
    'set default function folder
    path = "C:\WindTH"
    ret = SapModel.Options.SetDefaultFunctionFolder(path)
 
    'close SAP2000
    ret = SapObject.ApplicationExit(False)
 
    'clean up variables
    Set SapModel = Nothing
    Set SapObject = Nothing
    Set myHelper = Nothing
End Sub
```



## Release Notes

Initial release in version 25.0.0

## See Also

[GetDefaultFunctionFolder](SetDefaultFunctionFolder.md)

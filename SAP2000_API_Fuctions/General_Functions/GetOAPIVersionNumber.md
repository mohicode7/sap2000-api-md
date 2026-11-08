# GetOAPIVersionNumber

## Syntax

SapObject.GetOAPIVersionNumber

## VB6 Procedure

Function GetOAPIVersionNumber() As Double

## Parameters

None

## Remarks

Retrieves the
API version implemented by SAP2000. The returned API version can be compared
to Helper.GetOAPIVersionNumber to determine API compatibility with the
started/attached instance of SAP2000.

## VBA Example



```
Public Sub Example()
Dim myHelper As cHelper
Dim SapObject As cOAPI
Dim clientAPIVersion As Double
Dim programAPIVersion As Double
Dim ret As Integer
 
‘create SapObject
Set myHelper = New Helper
Set SapObject = myHelper.CreateObjectProgID("CSI.SAP2000.API.SapObject")
 
'start SAP2000 application
ret = SapObject.ApplicationStart()
 
'get client API version
clientAPIVersion = myHelper.GetOAPIVersionNumber()
 
'get program API version
programAPIVersion = SapObject.GetOAPIVersionNumber()
 
‘API compatibility check
If clientAPIVersion > programAPIVersion Then
    'API client uses a newer version 
 of the API than SAP2000.
    'All API functionality may be 
 not be available.
    Exit Sub
End If
 
'close SAP2000
ret = SapObject.ApplicationExit(False)
 
'clean up variables
Set SapObject = Nothingx
Set myHelper = Nothing
End Sub
 
```



## Release Notes

Initial release in version 17.00

## See Also

[Helper.GetOAPIVersionNumber](Helper/GetOAPIVersionNumber.md)

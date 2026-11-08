# GetVersion

## Syntax

SapObject.SapModel.GetVersion

## VB6 Procedure

Function GetVersion(ByRef Version As String, ByRef MyVersionNumber As Double) As Long

## Parameters

Version

The program version name that is externally displayed to the user.

MyVersionNumber

The program version number that is used internally by the program and not displayed to the user.

## Remarks

This function returns the SAP2000 program version.

The function returns zero if the information is successfully retrieved; otherwise it returns a nonzero value.

## VBA Example



```
Sub GetProgramVersion()

   'dimension variables

      Dim SapObject as cOAPI

      Dim SapModel As cSapModel

      Dim Version As String

      Dim MyVersionNumber As Double

      Dim ret As Long



   'create Sap2000 object

      Set SapObject = CreateObject("CSI.SAP2000.API.SapObject")



   'start Sap2000 application

      SapObject.ApplicationStart



   'create SapModel object

      Set SapModel = SapObject.SapModel



   'get program version

      ret = SapModel.GetVersion(Version,MyVersionNumber)



   'close Sap2000

      SapObject.ApplicationExit False

      Set SapModel = Nothing

      Set SapObject = Nothing

End Sub
```



## Release Notes

Initial release in version 12.01.

## See Also

# OpenFile

## Syntax

SapObject.SapModel.File.OpenFile

## VB6 Procedure

Function OpenFile(ByVal FileName As String) As Long

## Parameters

FileName

The full path of a model file to be opened in the Sap2000 application.

## Remarks

This function opens an existing Sap2000 file. The file name must have an sdb, $2k, s2k, xlsx, xls, or mdb extension. Files with sdb extensions are opened as standard Sap2000 files. Files with $2k and s2k extensions are imported as text files. Files with xlsx and xls extensions are imported as Microsoft Excel files. Files with mdb extensions are imported as Microsoft Access files.

This function returns zero if the file is successfully opened and nonzero if it is not opened.

The function is only applicable when you are accessing the Sap2000 API from an external application. It will return an error if you call it from VBA inside Sap2000.

## VBA Example



```
Sub OpenSDB()

   'dimension variables

      Dim SapObject as cOAPI

      Dim SapModel As cSapModel

      Dim FileName as String

      Dim ret as Long



   'create Sap2000 object

      Set SapObject = CreateObject("CSI.SAP2000.API.SapObject")



   'start Sap2000 application

      SapObject.ApplicationStart



   'create SapModel object

      Set SapModel = SapObject.SapModel



   'initialize model

      ret = SapModel.InitializeNewModel



   'open an existing file

      FileName = "C:\SapAPI\Example 1-019a.sdb"

      ret = SapModel.File.OpenFile(FileName)



   'close Sap2000

      SapObject.ApplicationExit False

      Set SapModel = Nothing

      Set SapObject = Nothing

End Sub
```



## Release Notes

Initial release in version 11.00.

Added reference to xlsx file extension in SAP2000 v15.0.0 and CSiBridge v15.1.0.

## See Also

[ApplicationStart](../General_Functions/ApplicationStart.md)

[Save](Save.md)

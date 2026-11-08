# ApplicationStart

## Syntax

SapObject.ApplicationStart

## VB6 Procedure

Function ApplicationStart(Optional ByVal Units As eUnits = kip\_in\_F, Optional ByVal Visible As Boolean = True, Optional ByVal FileName As String = "") As Long

## Parameters

Units

The database units used when a new model is created. Data is internally stored in the program in the database units. The database units may be one of the following items in the eUnits enumeration:

lb\_in\_F = 1

lb\_ft\_F = 2

kip\_in\_F = 3

kip\_ft\_F = 4

kN\_mm\_C = 5

kN\_m\_C = 6

kgf\_mm\_C = 7

kgf\_m\_C = 8

N\_mm\_C = 9

N\_m\_C = 10

Ton\_mm\_C = 11

Ton\_m\_C = 12

kN\_cm\_C = 13

kgf\_cm\_C = 14

N\_cm\_C = 15

Ton\_cm\_C = 16

Visible

If this item is True then the application is visible when started.  If it is False then the application is hidden when started.

FileName

The full path of a model file to be opened when the Sap2000 application is started. If no file name is specified, the application starts without loading an existing model.

## Remarks

This function starts the Sap2000 application.

When the model is not visible it does not appear on screen and it does not appear in the Windows task bar.

If no filename is specified, you can later open a model or create a model through the API.

The file name must have an .sdb, .$2k, .s2k, .xls, or .mdb extension. Files with .sdb extensions are opened as standard SAP2000 files. Files with .$2k and .s2k extensions are imported as text files. Files with .xls extensions are imported as Microsoft Excel files. Files with .mdb extensions are imported as Microsoft Access files.

This function returns zero if the application successfully starts and nonzero if it fails.

## VBA Example



```
Sub StartExample()

   'dimension variables

      Dim SapObject as cOAPI
      Dim ret as Long



   'create Sap2000 object

      Set SapObject = CreateObject("CSI.SAP2000.API.SapObject")



   'start Sap2000 application

      ret = SapObject.ApplicationStart



   'close Sap2000

      SapObject.ApplicationExit False

      Set SapObject = Nothing

End Sub
```



## Release Notes

Initial release in version 11.00.

## See Also

[ApplicationExit](SetAsActiveObject.md)

[OpenFile](../File/OpenFile.md)

[InitializeNewModel](InitializeNewModel.md)

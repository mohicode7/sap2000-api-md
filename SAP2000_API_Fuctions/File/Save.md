# Save

## Syntax

SapObject.SapModel.File.Save

## VB6 Procedure

Function FileSave(Optional ByVal FileName As String = "") As Long

## Parameters

FileName

The full path to which the model file is saved.

## Remarks

If a file name is specified, it should have an .sdb extension. If no file name is specified, the file is saved using its current name.

If there is no current name for the file (the file has not been saved previously) and this function is called with no file name specified, an error will be returned.

This function returns zero if the file is successfully saved and nonzero if it is not saved.

## VBA Example



```
Sub SaveSDB()

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



   'start a new template model

      ret = SapModel.File.New2DFrame(PortalFrame, 3, 124, 3, 200)



   'save SDB file

      ret=SapModel.File.Save("C:\SapAPI\x.sdb")



   'close Sap2000

      SapObject.ApplicationExit False

      Set SapModel = Nothing

      Set SapObject = Nothing

End Sub
```



## Release Notes

Initial release in version 11.00.

## See Also

[OpenFile](OpenFile.md)

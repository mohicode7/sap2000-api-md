# GetModelFilename

## Syntax

SapObject.SapModel.GetModelFilename

## VB6 Procedure

Function GetModelFilename(Optional ByVal IncludePath As Boolean = True) As String

## Parameters

IncludePath

A boolean (True or False) value. When this item is True, the returned filename includes the full path where the file is located.

## Remarks

The function returns a string that represents the filename of the current model, with or without the full path.

## VBA Example



```
Sub GetModelFilename()

   'dimension variables

      Dim SapObject as cOAPI

      Dim SapModel As cSapModel

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



   'save the model

      ret = SapModel.File.Save("C:\SapAPI\API_1-001.sdb")



   'display the filename of the model

      MsgBox “Model filename = “ & SapModel.GetModelFilename



   'close Sap2000

      SapObject.ApplicationExit False

      Set SapModel = Nothing

      Set SapObject = Nothing

End Sub
```



## Release Notes

Initial release in version 14.2.2

## See Also

[GetModelFilepath](GetModelFilepath.md)

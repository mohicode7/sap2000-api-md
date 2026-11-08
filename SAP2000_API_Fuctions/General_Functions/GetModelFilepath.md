# GetModelFilepath

## Syntax

SapObject.SapModel.GetModelFilepath

## VB6 Procedure

Function GetModelFilepath() As String

## Parameters

None

## Remarks

The function returns a string that represents the filepath of the current model.

## VBA Example



```
Sub GetModelFilepath()

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



   'display the path where the model is saved

      MsgBox “Model filepath = “ & SapModel.GetModelFilepath



   'close Sap2000

      SapObject.ApplicationExit False

      Set SapModel = Nothing

      Set SapObject = Nothing

End Sub
```



## Release Notes

Initial release in version 14.2.2

## See Also

[GetModelFilename](GetModelFilename.md)

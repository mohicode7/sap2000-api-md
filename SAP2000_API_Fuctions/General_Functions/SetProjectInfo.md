# SetProjectInfo

## Syntax

SapObject.SapModel.SetProjectInfo

## VB6 Procedure

Function SetProjectInfo(ByVal Item As String, ByVal Data As String) As Long

## Parameters

Item

The name of the project information item to be set.

Data

The data for the specified project information item.

## Remarks

This function sets the data for an item in the project information.

The function returns zero if the data is successfully set; otherwise it returns a nonzero value.

## VBA Example



```
Sub SetProjectInfoData()

   'dimension variables

      Dim SapObject as cOAPI

      Dim SapModel As cSapModel

      Dim ret As Long



   'create Sap2000 object

      Set SapObject = CreateObject("CSI.SAP2000.API.SapObject")



   'start Sap2000 application

      SapObject.ApplicationStart



   'create SapModel object

      Set SapModel = SapObject.SapModel



   'initialize model

      ret = SapModel.InitializeNewModel



   'create model from template

      ret = SapModel.File.New2DFrame(PortalFrame, 2, 144, 2, 288)



   'set project information data

      ret = SapModel.SetProjectInfo("Company Name", "Computers and Structures, Inc.")

      ret = SapModel.SetProjectInfo("Project Name", "API Testing")

      ret = SapModel.SetProjectInfo("My Item", "My Data")



   'close Sap2000

      SapObject.ApplicationExit False

      Set SapModel = Nothing

      Set SapObject = Nothing

End Sub
```



## Release Notes

Initial release in version 11.03.

## See Also

[GetUserComment](GetUserComment.md)

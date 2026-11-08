# GetProjectInfo

## Syntax

SapObject.SapModel.GetProjectInfo

## VB6 Procedure

Function GetProjectInfo(ByRef NumberItems As Long, ByRef Item() As String, ByRef Data() As String) As Long

## Parameters

NumberItems

The number of project info items returned.

Item

This is an array that includes the name of the project information item.

Data

This is an array that includes the data for the specified project information item.

## Remarks

This function retrieves the project information data.

The function returns zero if the data is successfully retrieved; otherwise it returns a nonzero value.

## VBA Example



```
Sub GetProjectInformationData()

   'dimension variables

      Dim SapObject as cOAPI

      Dim SapModel As cSapModel

      Dim ret As Long

      Dim NumberItems As Long

      Dim Item() As String

      Dim Data() As String



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



   'get project information data

      ret = SapModel.GetProjectInfo(NumberItems, Item, Data)



   'close Sap2000

      SapObject.ApplicationExit False

      Set SapModel = Nothing

      Set SapObject = Nothing

End Sub
```



## Release Notes

Initial release in version 11.03.

## See Also

[SetProjectInfo](SetProjectInfo.md)

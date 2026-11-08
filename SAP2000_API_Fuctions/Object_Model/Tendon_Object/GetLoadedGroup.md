# GetLoadedGroup

## Syntax

SapObject.SapModel.TendonObj.GetLoadedGroup

## VB6 Procedure

Function GetLoadedGroup(ByVal Name As String, ByRef GroupName As String) As Long

## Parameters

Name

The name of an existing tendon object.

GroupName

This is the name of an existing group. All objects in the specified group can be loaded by the tendon.

## Remarks

This function retrieves the loaded group for tendon objects. A tendon object transfers its load to any object that is in the specified group.

The function returns zero if the assignment is successfully retrieved, otherwise it returns a nonzero value.

## VBA Example



```
Sub GetTendonLoadedGroup()

   'dimension variables

      Dim SapObject as cOAPI

      Dim SapModel As cSapModel

      Dim ret As Long

      Dim GroupName As String

      Dim Name As String



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



   'add tendon object by points

      ret = SapModel.TendonObj.AddByPoint("3", "9", Name)



   'get tendon loaded group

      ret = SapModel.TendonObj.GetLoadedGroup(Name, GroupName)



   'close Sap2000

      SapObject.ApplicationExit False

      Set SapModel = Nothing

      Set SapObject = Nothing

End Sub
```



## Release Notes

Initial release in version 11.00.

## See Also

[SetLoadedGroup](SetLoadedGroup.md)

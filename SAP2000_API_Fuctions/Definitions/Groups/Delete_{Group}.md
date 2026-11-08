# Delete

## Syntax

SapObject.SapModel.GroupDef.Delete

## VB6 Procedure

Function Delete(ByVal Name As String) As Long

## Parameters

Name

The name of an existing group.

## Remarks

The function deletes the specified group. It will return an error if an attempt is made to delete the Group named ALL.

The function returns zero if the group is successfully deleted, otherwise it returns a nonzero value.

## VBA Example



```
Sub DeleteGroup()

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



   'create model from template

      ret = SapModel.File.New2DFrame(PortalFrame, 3, 124, 3, 200)



   'define new group

      ret = SapModel.GroupDef.SetGroup("Group1")



   'delete group

      ret = SapModel.GroupDef.Delete("Group1")



   'close Sap2000

      SapObject.ApplicationExit False

      Set SapModel = Nothing

      Set SapObject = Nothing

End Sub
```



## Release Notes

Initial release in version 11.00.

## See Also

[SetGroup](SetGroup_%7BGroup%7D.md)

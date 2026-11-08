# ChangeName

## Syntax

SapObject.SapModel.GroupDef.ChangeName

## VB6 Procedure

Function ChangeName(ByVal Name As String, ByVal NewName As String) As Long

## Parameters

Name

The existing name of a defined group.

NewName

The new name for the group.

## Remarks

The function returns zero if the new name is successfully applied, otherwise it returns a nonzero value.

Changing the name of group ALL will fail and return an error.

## VBA Example



```
Sub ChangeGroupName()

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



   'change name of group

      ret = SapModel.GroupDef.ChangeName("Group1", "MyGroup")



   'close Sap2000

      SapObject.ApplicationExit False

      Set SapModel = Nothing

      Set SapObject = Nothing

End Sub
```



## Release Notes

Initial release in version 11.00.

## See Also

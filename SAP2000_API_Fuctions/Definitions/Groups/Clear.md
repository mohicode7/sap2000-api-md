# Clear

## Syntax

SapObject.SapModel.GroupDef.Clear

## VB6 Procedure

Function Clear(ByVal Name As String) As Long

## Parameters

Name

The name of an existing group.

## Remarks

This function clears (removes) all assignments from the specified group.

The function returns zero if the group assignment is successfully cleared, otherwise it returns a nonzero value.

## VBA Example



```
Sub ClearGroup()

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



   'define new group

      ret = SapModel.GroupDef.SetGroup("Group1")



   'add frame objects to group

      ret = SapModel.FrameObj.SetGroupAssign("8", "Group1")

      ret = SapModel.FrameObj.SetGroupAssign("10", "Group1")



   'clear group

      ret = SapModel.GroupDef.Clear("Group1")



   'close Sap2000

      SapObject.ApplicationExit False

      Set SapModel = Nothing

      Set SapObject = Nothing

End Sub
```



## Release Notes

Initial release in version 11.00.

## See Also

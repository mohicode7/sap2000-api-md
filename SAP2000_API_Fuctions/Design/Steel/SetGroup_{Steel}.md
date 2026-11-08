# SetGroup

## Syntax

SapObject.SapModel.DesignSteel.SetGroup

## VB6 Procedure

Function SetGroup(ByVal Name As String, ByVal Selected As Boolean) As Long

## Parameters

Name

The name of an existing group.

Selected

If this item is True, the specified group is selected as a design group for steel design. If it is False, the group is not selected for steel design.

## Remarks

This function selects or deselects a group for steel design.

The function returns zero if the selection status is successfully set; otherwise it returns a nonzero value.

## VBA Example



```
Sub SetSteelDesignGroup()

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



   'select group for steel design

      ret = SapModel.DesignSteel.SetGroup("ALL", True)



   'close Sap2000

      SapObject.ApplicationExit False

      Set SapModel = Nothing

      Set SapObject = Nothing

End Sub
```



## Release Notes

Initial release in version 11.03.

## See Also

[GetGroup](GetGroup_%7BSteel%7D.md)

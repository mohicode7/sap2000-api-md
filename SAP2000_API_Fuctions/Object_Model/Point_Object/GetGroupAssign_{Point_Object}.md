# GetGroupAssign

## Syntax

SapObject.SapModel.PointObj.GetGroupAssign

## VB6 Procedure

Function GetGroupAssign(ByVal Name As String, ByRef NumberGroups As Long, ByRef Groups() As String)  As Long

## Parameters

Name

The name of an existing point object.

NumberGroups

The number of group names retrieved.

Groups

The names of the groups to which the point object is assigned.

## Remarks

This function retrieves the names of the groups to which a specified point object is assigned.

The function returns zero if the group assignments are successfully retrieved, otherwise it returns a nonzero value.

## VBA Example



```
Sub GetPointObjectGroups()

   'dimension variables

      Dim SapObject as cOAPI

      Dim SapModel As cSapModel

      Dim ret As Long

      Dim NumberGroups As Long

      Dim Groups() As String



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



   'define new groups

      ret = SapModel.GroupDef.SetGroup("Group1")

      ret = SapModel.GroupDef.SetGroup("Group2")



   'add point object to groups

      ret = SapModel.PointObj.SetGroupAssign("3", "Group1")

      ret = SapModel.PointObj.SetGroupAssign("3", "Group2")



   'get point object groups

      ret = SapModel.PointObj.GetGroupAssign("3", NumberGroups, Groups)



   'close Sap2000

      SapObject.ApplicationExit False

      Set SapModel = Nothing

      Set SapObject = Nothing

End Sub
```



## Release Notes

Initial release in version 20.0.0.

## See Also

[SetGroupAssign](SetGroupAssign_%7BPoint_Object%7D.md)

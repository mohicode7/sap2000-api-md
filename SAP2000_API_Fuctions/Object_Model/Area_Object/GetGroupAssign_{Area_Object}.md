# GetGroupAssign

## Syntax

SapObject.SapModel.AreaObj.GetGroupAssign

## VB6 Procedure

Function GetGroupAssign(ByVal Name As String, ByRef NumberGroups As Long, ByRef Groups() As String)  As Long

## Parameters

Name

The name of an existing area object.

NumberGroups

The number of group names retrieved.

Groups

The names of the groups to which the area object is assigned.

## Remarks

This function retrieves the names of the groups to which a specified area object is assigned.

The function returns zero if the group assignments are successfully retrieved, otherwise it returns a nonzero value.

## VBA Example



```
Sub GetAreaObjectGroups()

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

      ret = SapModel.File.NewWall(2, 48, 2, 48)



   'define new groups

      ret = SapModel.GroupDef.SetGroup("Group1")

      ret = SapModel.GroupDef.SetGroup("Group2")



   'add area object to groups

      ret = SapModel.AreaObj.SetGroupAssign("1", "Group1")

      ret = SapModel.AreaObj.SetGroupAssign("1", "Group2")



   'get area object groups

      ret = SapModel.AreaObj.GetGroupAssign("1", NumberGroups, Groups)



   'close Sap2000

      SapObject.ApplicationExit False

      Set SapModel = Nothing

      Set SapObject = Nothing

End Sub
```



## Release Notes

Initial release in version 20.0.0.

## See Also

[SetGroupAssign](SetGroupAssign_%7BArea_Object%7D.md)

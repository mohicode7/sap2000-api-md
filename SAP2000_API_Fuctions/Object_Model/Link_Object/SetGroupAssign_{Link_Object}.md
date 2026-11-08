# SetGroupAssign

## Syntax

SapObject.SapModel.LinkObj.SetGroupAssign

## VB6 Procedure

Function SetGroupAssign(ByVal Name As String, ByVal GroupName As String, Optional By Val Remove As Boolean = False, Optional ByVal ItemType As eItemType = Object) As Long

## Parameters

Name

The name of an existing link object or group, depending on the value of the ItemType item.

GroupName

The name of an existing group to which the assignment is made.

Remove

If this item is False, the specified link objects are added to the group specified by the GroupName item. If it is True, the link objects are removed from the group.

ItemType

This is one of the following items in the eItemType enumeration:

Object = 0

Group = 1

SelectedObjects = 2

If this item is Object, the link object specified by the Name item is added or removed from the group specified by the GroupName item.

If this item is Group, all link objects in the group specified by the Name item are added or removed from the group specified by the GroupName item.

If this item is SelectedObjects, all selected link objects are added or removed from the group specified by the GroupName item, and the Name item is ignored.

## Remarks

This function adds or removes link objects from a specified group.

The function returns zero if the group assignment is successful; otherwise it returns a nonzero value.

## VBA Example



```
Sub AddLinkObjectsToGroup()

   'dimension variables

      Dim SapObject as cOAPI

      Dim SapModel As cSapModel

      Dim ret As Long

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



   'add link object by points

      ret = SapModel.LinkObj.AddByPoint("1", "5", Name)



   'define new group

      ret = SapModel.GroupDef.SetGroup("Group1")



   'add link object to group

      ret = SapModel.LinkObj.SetGroupAssign(Name, "Group1")



   'select objects in group

      ret = SapModel.SelectObj.Group("Group1")



   'close Sap2000

      SapObject.ApplicationExit False

      Set SapModel = Nothing

      Set SapObject = Nothing

End Sub
```



## Release Notes

Initial release in version 11.00.

## See Also

[GetGroupAssign](GetGroupAssign_%7BLink_Object%7D.md)

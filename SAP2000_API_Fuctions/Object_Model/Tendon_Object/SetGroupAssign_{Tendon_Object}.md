# SetGroupAssign

## Syntax

SapObject.SapModel.TendonObj.SetGroupAssign

## VB6 Procedure

Function SetGroupAssign(ByVal Name As String, ByVal GroupName As String, Optional By Val Remove As Boolean = False, Optional ByVal ItemType As eItemType = Object) As Long

## Parameters

Name

The name of an existing tendon object or group, depending on the value of the ItemType item.

GroupName

The name of an existing group to which the assignment is made.

Remove

If this item is False, the specified tendon objects are added to the group specified by the GroupName item. If it is True, the tendon objects are removed from the group.

ItemType

This is one of the following items in the eItemType enumeration:

Object = 0

Group = 1

SelectedObjects = 2

If this item is Object, the tendon object specified by the Name item is added or removed from the group specified by the GroupName item.

If this item is Group, all tendon objects in the group specified by the Name item are added or removed from the group specified by the GroupName item.

If this item is SelectedObjects, all selected tendon objects are added or removed from the group specified by the GroupName item, and the Name item is ignored.

## Remarks

This function adds or removes tendon objects from a specified group.

The function returns zero if the group assignment is successful, otherwise it returns a nonzero value.

## VBA Example



```
Sub AddTendonObjectsToGroup()

   'dimension variables

      Dim SapObject as cOAPI

      Dim SapModel As cSapModel

      Dim ret As Long

      Dim Name1 As String

      Dim Name2 As String

      Dim NumberPoints As Long

      Dim MyType() As Long

      Dim x() As Double

      Dim y() As Double

      Dim z() As Double



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

      ret = SapModel.TendonObj.AddByPoint("3", "9", Name1)

      ret = SapModel.TendonObj.AddByPoint("2", "8", Name2)



   'set tendon geometry

      NumberPoints = 3

      ReDim MyType(NumberPoints - 1)

      ReDim x(NumberPoints - 1)

      ReDim y(NumberPoints - 1)

      ReDim z(NumberPoints - 1)

      MyType(0) = 1

      MyType(1) = 7

      MyType(2) = 6

      x(0) = 0:  y(0) = 0

      x(1) = 288:  y(1) = -12

      x(2) = 576:  y(2) = 0

      ret = SapModel.TendonObj.SetTendonData(Name1, NumberPoints, MyType, x, y, z, "Local")

      ret = SapModel.TendonObj.SetTendonData(Name2, NumberPoints, MyType, x, y, z, "Local")



   'define new group

      ret = SapModel.GroupDef.SetGroup("Group1")



   'add tendon objects to group

      ret = SapModel.TendonObj.SetGroupAssign(Name1, "Group1")

      ret = SapModel.TendonObj.SetGroupAssign(Name2, "Group1")



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

[GetGroupAssign](GetGroupAssign_%7BTendon_Object%7D.md)

# GetAssignments

## Syntax

SapObject.SapModel.GroupDef.GetAssignments

## VB6 Procedure

Function GetAssignments(ByVal Name As String, ByRef NumberItems As Long, ByRef ObjectType() As Long, ByRef ObjectName() As String) As Long

## Parameters

Name

The name of an existing group.

NumberItems

The number of assignments made to the specified group.

ObjectType

This is an array that includes the object type of each item in the group.

1 = Point object

2 = Frame object

3 = Cable object

4 = Tendon object

5 = Area object

6 = Solid object

7 = Link object

ObjectName

This is an array that includes the name of each item in the group.

## Remarks

This function retrieves the assignments to a specified group.

The function returns zero if the group assignment is successfully retrieved, otherwise it returns a nonzero value.

## VBA Example



```
Sub GetGroupAssignments()

   'dimension variables

      Dim SapObject as cOAPI

      Dim SapModel As cSapModel

      Dim ret As Long

      Dim Name As String

      Dim NumberItems As Long

      Dim ObjectType() As Long

      Dim ObjectName() As String



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



   'add cable object by points

      ret = SapModel.CableObj.AddByPoint("1", "5", Name)



   'set cable data

      ret = SapModel.CableObj.SetCableData(Name, 7, 1, 0, 0, 24)



   'define new group

      ret = SapModel.GroupDef.SetGroup("Group1")



   'add point objects to group

      ret = SapModel.PointObj.SetGroupAssign("3", "Group1")

      ret = SapModel.PointObj.SetGroupAssign("6", "Group1")

      ret = SapModel.PointObj.SetGroupAssign("9", "Group1")



   'add frame objects to group

      ret = SapModel.FrameObj.SetGroupAssign("8", "Group1")

      ret = SapModel.FrameObj.SetGroupAssign("10", "Group1")



   'add cable object to group

      ret = SapModel.CableObj.SetGroupAssign(Name, "Group1")



   'get group assignments

      ret = SapModel.GroupDef.GetAssignments("Group1", NumberItems, ObjectType, ObjectName)



   'close Sap2000

      SapObject.ApplicationExit False

      Set SapModel = Nothing

      Set SapObject = Nothing

End Sub
```



## Release Notes

Initial release in version 11.00.

## See Also

# GetGroupAssign

## Syntax

SapObject.SapModel.LinkObj.GetGroupAssign

## VB6 Procedure

Function GetGroupAssign(ByVal Name As String, ByRef NumberGroups As Long, ByRef Groups() As String)  As Long

## Parameters

Name

The name of an existing link object.

NumberGroups

The number of group names retrieved.

Groups

The names of the groups to which the link object is assigned.

## Remarks

This function retrieves the names of the groups to which a specified link object is assigned.

The function returns zero if the group assignments are successfully retrieved, otherwise it returns a nonzero value.

## VBA Example



```
Sub GetLinkObjectGroups()

   'dimension variables

      Dim SapObject as cOAPI

      Dim SapModel As cSapModel

      Dim ret As Long
      Dim Name As String

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
 
   'add link object by points

      ret = SapModel.LinkObj.AddByPoint("1", "5", Name)



   'define new groups

      ret = SapModel.GroupDef.SetGroup("Group1")

      ret = SapModel.GroupDef.SetGroup("Group2")



   'add link object to groups

      ret = SapModel.LinkObj.SetGroupAssign(Name, "Group1")

      ret = SapModel.LinkObj.SetGroupAssign(Name, "Group2")



   'get link object groups

      ret = SapModel.LinkObj.GetGroupAssign(Name, NumberGroups, Groups)



   'close Sap2000

      SapObject.ApplicationExit False

      Set SapModel = Nothing

      Set SapObject = Nothing

End Sub
```



## Release Notes

Initial release in version 20.0.0.

## See Also

[SetGroupAssign](SetGroupAssign_%7BLink_Object%7D.md)

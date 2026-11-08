# GetGroupAssign

## Syntax

SapObject.SapModel.CableObj.GetGroupAssign

## VB6 Procedure

Function GetGroupAssign(ByVal Name As String, ByRef NumberGroups As Long, ByRef Groups() As String)  As Long

## Parameters

Name

The name of an existing cable object.

NumberGroups

The number of group names retrieved.

Groups

The names of the groups to which the cable object is assigned.

## Remarks

This function retrieves the names of the groups to which a specified cable object is assigned.

The function returns zero if the group assignments are successfully retrieved, otherwise it returns a nonzero value.

## VBA Example



```
Sub GetCableObjectGroups()

   'dimension variables

      Dim SapObject as cOAPI

      Dim SapModel As cSapModel

      Dim ret As Long

      Dim Name1 As String

      Dim Name2 As String

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

      ret = SapModel.File.New2DFrame(PortalFrame, 3, 124, 3, 200)



   'add cable object by points

      ret = SapModel.CableObj.AddByPoint("1", "6", Name1)

      ret = SapModel.CableObj.AddByPoint("5", "10", Name2)
   'set cable data

      ret = SapModel.CableObj.SetCableData(Name1, 7, 1, 0, 0, 24)

      ret = SapModel.CableObj.SetCableData(Name2, 7, 1, 0, 0, 24)



   'define new groups

      ret = SapModel.GroupDef.SetGroup("Group1")

      ret = SapModel.GroupDef.SetGroup("Group2")



   'add cable object to groups

      ret = SapModel.CableObj.SetGroupAssign(Name1, "Group1")

      ret = SapModel.CableObj.SetGroupAssign(Name1, "Group2")



   'get cable object groups

      ret = SapModel.CableObj.GetGroupAssign(Name1, NumberGroups, Groups)



   'close Sap2000

      SapObject.ApplicationExit False

      Set SapModel = Nothing

      Set SapObject = Nothing

End Sub
```



## Release Notes

Initial release in version 20.0.0.

## See Also

[SetGroupAssign](SetGroupAssign_%7BCable_Object%7D.md)

# GetGroupAssign

## Syntax

SapObject.SapModel.TendonObj.GetGroupAssign

## VB6 Procedure

Function GetGroupAssign(ByVal Name As String, ByRef NumberGroups As Long, ByRef Groups() As String) ) As Long

## Parameters

Name

The name of an existing tendon object.

NumberGroups

The number of group names retrieved.

Groups

The names of the groups to which the tendon object is assigned.

## Remarks

This function retrieves the names of the groups to which a specified tendon object is assigned.

The function returns zero if the group assignments are successfully retrieved, otherwise it returns a nonzero value.

## VBA Example



```
Sub GetTendonObjectGroups()

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



   'define new groups

      ret = SapModel.GroupDef.SetGroup("Group1")
      ret = SapModel.GroupDef.SetGroup("Group2")



   'add tendon object to groups

      ret = SapModel.TendonObj.SetGroupAssign(Name2, "Group1")

      ret = SapModel.TendonObj.SetGroupAssign(Name2, "Group2")



   'get tendon object groups

      ret = SapModel.TendonObj.GetGroupAssign(Name2, NumberGroups, Groups)



   'close Sap2000

      SapObject.ApplicationExit False

      Set SapModel = Nothing

      Set SapObject = Nothing

End Sub
```



## Release Notes

Initial release in v. 20.0.0.

## See Also

[SetGroupAssign](SetGroupAssign_%7BTendon_Object%7D.md)

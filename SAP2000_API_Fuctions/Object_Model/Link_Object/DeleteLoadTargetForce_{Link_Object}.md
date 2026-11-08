# DeleteLoadTargetForce

## Syntax

SapObject.SapModel.LinkObj.DeleteLoadTargetForce

## VB6 Procedure

Function DeleteLoadTargetForce(ByVal Name As String, ByVal LoadPat As String, Optional ByVal ItemType As eItemType = Object) As Long

## Parameters

Name

The name of an existing link object or group, depending on the value of the ItemType item.

LoadPat

The name of a defined load pattern.

ItemType

This is one of the following items in the eItemType enumeration:

Object = 0

Group = 1

SelectedObjects = 2

If this item is Object, the load assignments are deleted for the link object specified by the Name item.

If this item is Group, the load assignments are deleted for all link objects in the group specified by the Name item.

If this item is SelectedObjects, the load assignments are deleted for all selected link objects, and the Name item is ignored.

## Remarks

This function deletes the target force assignments to the specified link objects for the specified load pattern.

The function returns zero if the target force assignments are successfully deleted; otherwise it returns a nonzero value.

## VBA Example



```
Sub DeleteLinkTargetForce()

   'dimension variables

      Dim SapObject as cOAPI

      Dim SapModel As cSapModel

      Dim ret As Long

      Dim DOF() As Boolean

      Dim f() As double

      Dim RD() As double

      Dim Name1 As String

      Dim Name2 As String



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

      ret = SapModel.LinkObj.AddByPoint("1", "5", Name1)

      ret = SapModel.LinkObj.AddByPoint("2", "6", Name2)



   'assign link target force

      ReDim DOF(5)

      ReDim f(5)

      ReDim RD(5)

      DOF(0) = True

      f(0) = 50

      RD(0) = 1

      ret = SapModel.LinkObj.SetLoadTargetForce("ALL", "DEAD", DOF, f, RD, Group)



   'delete link target force

      ret = SapModel.LinkObj.DeleteLoadTargetForce(Name1, "DEAD")



   'close Sap2000

      SapObject.ApplicationExit False

      Set SapModel = Nothing

      Set SapObject = Nothing

End Sub
```



## Release Notes

Initial release in version 11.00.

Changed nomenclature from Load Cases, Analysis Cases and Response Combinations to Load Patterns, Load Cases and Load Combinations, respectively, in version 12.00.

## See Also

[GetLoadTargetForce](GetLoadTargetForce_%7BLink_Object%7D.md)

[SetLoadTargetForce](SetLoadTargetForce_%7BLink_Object%7D.md)

# GetLoadGravity

## Syntax

SapObject.SapModel.LinkObj.GetLoadGravity

## VB6 Procedure

Function GetLoadGravity(ByVal Name As String, ByRef NumberItems As Long, ByRef LinkName() As String, ByRef LoadPat() As String, ByRef CSys() As String, ByRef x() As Double, ByRef y() As Double, ByRef z() As Double, Optional ByVal ItemType As eItemType = Object) As Long

## Parameters

Name

The name of an existing link object or group, depending on the value of the ItemType item.

NumberItems

The total number of gravity loads retrieved for the specified link objects.

LinkName

This is an array that includes the name of the link object associated with each gravity load.

LoadPat

This is an array that includes the name of the coordinate system in which the gravity load multipliers are specified.

CSys

This is an array that includes the name of the coordinate system associated with each gravity load.

x, y, z

These are arrays of gravity load multipliers in the x, y and z directions of the specified coordinate system.

ItemType

This is one of the following items in the eItemType enumeration:

Object = 0

Group = 1

SelectedObjects = 2

If this item is Object, the assignments are retrieved for the link object specified by the Name item.

If this item is Group, the assignments are retrieved for all link objects in the group specified by the Name item.

If this item is SelectedObjects, assignments are retrieved for all selected link objects, and the Name item is ignored.

## Remarks

This function retrieves the gravity load assignments to link objects.

The function returns zero if the load assignments are successfully retrieved; otherwise it returns a nonzero value.

## VBA Example



```
Sub GetLinkGravityLoad()

   'dimension variables

      Dim SapObject as cOAPI

      Dim SapModel As cSapModel

      Dim ret As Long

      Dim NumberItems As Long

      Dim LinkName() As String

      Dim LoadPat() As String

      Dim CSys() As String

      Dim x() As Double

      Dim y() As Double

      Dim z() As Double

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



   'assign link gravity loads

      ret = SapModel.LinkObj.SetLoadGravity("ALL", "DEAD", 0, 0, -1, , , Group)



   'get link gravity load

      ret = SapModel.LinkObj.GetLoadGravity(Name, NumberItems, LinkName, LoadPat, CSys, x, y, z)



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

[SetLoadGravity](SetLoadGravity_%7BLink_Object%7D.md)

[DeleteLoadGravity](DeleteLoadGravity_%7BLink_Object%7D.md)

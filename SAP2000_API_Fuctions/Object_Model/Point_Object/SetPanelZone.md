# SetPanelZone

## Syntax

SapObject.SapMdel.PointObj.SetPanelZone

## VB6 Procedure

Function SetPanelZone(ByVal Name As String, ByVal PropType As Long, ByVal Thickness As Double, ByVal K1 As Double, ByVal K2 As Double, ByVal LinkProp As String, ByVal Connectivity As Long, ByVal LocalAxisFrom As Long, ByVal LocalAxisAngle As Double, Optional ByVal ItemType As eItemType = Object) As Long

## Parameters

Name

The name of an existing point object.

PropType

This is 0, 1, 2, or 3.

0 = Properties are elastic from column

1 = Properties are elastic from column and doubler plate

2 = Properties are from specified spring stiffnesses

3 = Properties are from a specified link property

Thickness

The thickness of the doubler plate. This item applies only when PropType = 1. [L]

K1

The spring stiffness for major axis bending (about the local 3 axis of the column and panel zone). This item applies only when PropType = 2. [FL/rad]

K2

The spring stiffness for minor axis bending (about the local 2 axis of the column and panel zone). This item applies only when PropType = 2. [FL/rad]

LinkProp

The name of the link property used to define the panel zone. This item applies only when PropType = 3.

Connectivity

This is 0 or 1.

0 = Panel zone connects beams to other objects

1 = Panel zone connects braces to other objects

LocalAxisFrom

This is 0 or 1.

0 = Panel zone local axis angle is from column

1 = Panel zone local axis angle is user defined

The LocalAxisFrom item can be 1 only when the PropType item is 3.

LocalAxisAngle

This item applies only when PropType = 3 and LocalAxisFrom = 1. It is the angle measured counter clockwise from the positive global X-axis to the local 2-axis of the panel zone. [deg]

ItemType

This is one of the following items in the eItemType enumeration:

Object = 0

Group = 1

SelectedObjects = 2

If this item is Object, the panel zone assignment is made to the point object specified by the Name item.

If this item is Group, the panel zone assignment is made to all point objects in the group specified by the Name item.

If this item is SelectedObjects, the panel zone assignment is made to all selected point objects and the Name item is ignored.

## Remarks

This function makes panel zone assignments to point objects. Any existing panel zone assignments are replaced by the new assignments.

The function returns zero if the panel zone data is successfully assigned, otherwise it returns a nonzero value.

## VBA Example



```
Sub AssignPanelZoneData()

   'dimension variables

      Dim SapObject as cOAPI

      Dim SapModel As cSapModel

      Dim ret As Long



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



   'assign panel zone

      ret = SapModel.PointObj.SetPanelZone("3", 1, 2, 0, 0, "", 0, 0, 0)



   'close Sap2000

      SapObject.ApplicationExit False

      Set SapModel = Nothing

      Set SapObject = Nothing

End Sub
```



## Release Notes

Initial release in version 11.00.

## See Also

[GetPanelZone](GetPanelZone.md)

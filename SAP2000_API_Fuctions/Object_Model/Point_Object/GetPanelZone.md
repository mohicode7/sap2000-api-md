# GetPanelZone

## Syntax

SapObject.SapModel.PointObj.GetPanelZone

## VB6 Procedure

Function GetPanelZone(ByVal Name As String, ByRef PropType As Long, ByRef Thickness As Double, ByRef K1 As Double, ByRef K2 As Double, ByRef LinkProp As String, ByRef Connectivity As Long, ByRef LocalAxisFrom As Long, ByRef LocalAxisAngle As Double) As Long

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

## Remarks

This function retrieves the panel zone assignment data for a point object.

The function returns zero if the panel zone data is successfully retrieved, otherwise it returns a nonzero value.

If no panel zone assignment is made to the point object, an error is returned.

## VBA Example



```
Sub GetPanelZoneData()

   'dimension variables

      Dim SapObject as cOAPI

      Dim SapModel As cSapModel

      Dim ret As Long

      Dim PropType As Long

      Dim Thickness As Double

      Dim K1 As Double

      Dim K2 As Double

      Dim LinkProp As String

      Dim Connectivity As Long

      Dim LocalAxisFrom As Long

      Dim LocalAxisAngle As Double



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



   'get panel zone data

      ret = SapModel.PointObj.GetPanelZone("3", PropType, Thickness, K1, K2, LinkProp, Connectivity, LocalAxisFrom, LocalAxisAngle)



   'close Sap2000

      SapObject.ApplicationExit False

      Set SapModel = Nothing

      Set SapObject = Nothing

End Sub
```



## Release Notes

Initial release in version 11.00.

## See Also

[SetPanelZone](SetPanelZone.md)

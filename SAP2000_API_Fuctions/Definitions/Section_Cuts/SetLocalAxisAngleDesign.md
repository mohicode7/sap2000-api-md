# SetLocalAxisAngleDesign

## Syntax

SapObject.SapModel.SectCut.SetLocalAxesAngleDesign

## VB6 Procedure

Function SetLocalAxesAngleDesign(ByVal Name As String, ByVal Angle As Double) As Long

## Parameters

Name

The name of an existing section cut.

Angle

For design local axes orientation type wall this is the angle from the global X to the local 2 axis.  For orientation types spandrel and slab it is the angle from the global X to the local 1 axis.

## Remarks

This function sets the local axes angle for section cuts whose result type is Design (Wall, Spandrel or Slab).

The function returns zero if the angle is successfully set, otherwise it returns a nonzero value.

## VBA Example

Sub SetSectionCutLocalAxesAngleDesign ()

'dimension variables

Dim SapObject As Sap2000v16.SapObject

Dim SapModel As cSapModel

Dim ret As Long

'create Sap2000 object

Set SapObject = New Sap2000v16.SapObject

'start Sap2000 application

SapObject.ApplicationStart

'create SapModel object

Set SapModel = SapObject.SapModel

'initialize model

ret = SapModel.InitializeNewModel

'create model from template

ret = SapModel.File.New2DFrame(PortalFrame, 2, 144, 2, 288)

'define new group

ret = SapModel.GroupDef.SetGroup("Group1")

'add objects to group

ret = SapModel.PointObj.SetGroupAssign("1", "Group1")

ret = SapModel.FrameObj.SetGroupAssign("1", "Group1")

'define section cut with result type of Design - Wall

ret = SapModel.SectCut.SetByGroup("SCut1", "Group1", 2)

'set section cut local axes angle

ret = SapModel.SectCut.SetLocalAxesAngleDesign("SCut1", 24)

'close Sap2000

SapObject.ApplicationExit False

Set SapModel = Nothing

Set SapObject = Nothing

End Sub

## Release Notes

Initial release in version 16.0.0.

## See Also

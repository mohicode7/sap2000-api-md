# SetByGroup

## Syntax

SapObject.SapModel.SectCut.SetByGroup

## VB6 Procedure

Function SetByGroup(ByVal Name As String, ByVal GroupName As String, ByVal MyType As Long) As Long

## Parameters

Name

The section cut name. If a section cut with this name already exists then the section cut is reinitialized with the new data.  All previous data assigned to the section cut is lost. If a section cut with this name does not exist then a new section cut is added.

GroupName

The name of the group on which the section cut is based.

MyType

This is either 1, 2, 3 or 4 indicating the result type of the section cut.

1 = Analysis

2 = Design Wall

3 = Design Spandrel

4 = Design Slab

## Remarks

This function adds a new section cut defined by a group to the model or reinitializes an existing section cut to be defined by a group.

The function returns zero if the section cut is successfully added or initialized, otherwise it returns a nonzero value.

## VBA Example

Sub SetSectionCutByGroup()

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

'define section cut

ret = SapModel.SectCut.SetByGroup("SCut1", "Group1", 1)

'close Sap2000

SapObject.ApplicationExit False

Set SapModel = Nothing

Set SapObject = Nothing

End Sub

## Release Notes

Initial release in version 16.0.0.

## See Also

[SetByQuad](SetByQuad.md)

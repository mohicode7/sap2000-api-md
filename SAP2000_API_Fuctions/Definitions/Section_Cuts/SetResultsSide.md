# SetResultsSide

## Syntax

SapObject.SapModel.SectCut.SetResultsSide

## VB6 Procedure

Function SetResultsSide(ByVal Name As String, ByVal Side As Long) As Long

## Parameters

Name

The name of an existing section cut.

Side

This item is either 1 or 2 and indicates the side of the elements from which section cut results are obtained.

For section cuts defined from quadrilaterals with an Analysis result type:

1 = Positive 3-axis side of quadrilateral

2 = Negative 3-axis side of quadrilateral

For section cuts with a Design Wall result type:

1 = Top

2 = Bottom

For section cuts with a Design Spandrel or Design Slab result type:

1 = Right

2 = Left

## Remarks

This function sets the side of the elements from which results are obtained.

The function returns zero if the side is successfully set, otherwise it returns a nonzero value.

## VBA Example

Sub SetSectionCutResultsSide()

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

'set section cut results side

ret = SapModel.SectCut.SetResultsSide ("SCut1", 2)

'close Sap2000

SapObject.ApplicationExit False

Set SapModel = Nothing

Set SapObject = Nothing

End Sub

## Release Notes

Initial release in version 16.0.0.

## See Also

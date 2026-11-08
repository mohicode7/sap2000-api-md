# SetByQuad

## Syntax

SapObject.SapModel.SectCut.SetByQuad

## VB6 Procedure

Function SetByQuad(ByVal Name As String, ByVal GroupName As String, ByVal MyType As Long, ByRef X() As Double, ByRef Y() As Double, ByRef Z() As Double) As Long

## Parameters

Name

The section cut name. If a section cut with this name already exists then the section cut is reinitialized with the new data.  All previous data assigned to the section cut is lost. If a section cut with this name does not exist then a new section cut is added.

GroupName

The name of the group associated with the section cut.

MyType

This is either 1, 2, 3 or 4 indicating the result type of the section cut.

1 = Analysis

2 = Design Wall

3 = Design Spandrel

4 = Design Slab

X

This is an array of four X coordinates, one for each of the four points defining the quadrilateral.

Y

This is an array of four Y coordinates, one for each of the four points defining the quadrilateral.

Z

This is an array of four Z coordinates, one for each of the four points defining the quadrilateral.

## Remarks

This function adds a new section cut defined by a quadrilateral to the model or reinitializes an existing section cut to be defined by a quadrilateral.

The function returns zero if the section cut is successfully added or initialized, otherwise it returns a nonzero value.

## VBA Example

Sub SetSectionCutByQuad()

'dimension variables

Dim SapObject As Sap2000v16.SapObject

Dim SapModel As cSapModel

Dim ret As Long

Dim X(3) As Double

Dim Y(3) As Double

Dim Z(3) As Double

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

ret = SapModel.FrameObj.SetGroupAssign("1", "Group1")

ret = SapModel.FrameObj.SetGroupAssign("2", "Group1")

ret = SapModel.FrameObj.SetGroupAssign("3", "Group1")

'define section cut

X(0) = -300

X(1) = 300

X(2) = 300

X(3) = -300

Y(0) = 10

Y(1) = 10

Y(2) = -10

Y(3) = -10

Z(0) = 10

Z(1) = 10

Z(2) = 10

Z(3) = 10

ret = SapModel.SectCut.SetByQuad("SCut1", "Group1", 1, X, Y, Z)

'close Sap2000

SapObject.ApplicationExit False

Set SapModel = Nothing

Set SapObject = Nothing

End Sub

## Release Notes

Initial release in version 16.0.0.

## See Also

[SetByGroup](SetByGroup.md)

[AddQuad](AddQuad.md)

[DeleteCut](DeleteCut.md)

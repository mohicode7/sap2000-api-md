# GetQuad

## Syntax

SapObject.SapModel.SectCut.GetQuad

## VB6 Procedure

Function AddQuad(ByVal Name As String, ByVal Num As Long, ByRef X() As Double, ByRef Y() As Double, ByRef Z() As Double) As Long

## Parameters

Name

The name of an existing section cut defined using quadrilateral cutting planes.

Num

The number of a quadirilateral cutting plane in the section cut.

GroupName

X

This is an array of four X coordinates, one for each of the four points defining the quadrilateral.

Y

This is an array of four Y coordinates, one for each of the four points defining the quadrilateral.

Z

This is an array of four Z coordinates, one for each of the four points defining the quadrilateral.

## Remarks

This function returns the coordinates of a quadrilateral cutting plane in a section cut defined by quadrilaterals.

The function returns zero if the coordinates are successfully returned, otherwise it returns a nonzero value.

## VBA Example

Sub GetSectionCutQuad()

'dimension variables

Dim SapObject As Sap2000v16.SapObject

Dim SapModel As cSapModel

Dim ret As Long

Dim X() As Double

Dim Y() As Double

Dim Z() As Double

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

Redim X(3)

Redim Y(3)

Redim Z(3)

X(0) = -300

X(1) = -270

X(2) = -270

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

'add a second quadrilateral

X(0) = 270

X(1) = 300

X(2) = 300

X(3) = 270

Y(0) = 10

Y(1) = 10

Y(2) = -10

Y(3) = -10

Z(0) = 10

Z(1) = 10

Z(2) = 10

Z(3) = 10

ret = SapModel.SectCut.AddQuad("SCut1", X, Y, Z)

'get coordinates of quadrilateral 2

ret = SapModel.SectCut.GetQuad("SCut1", 2, X, Y, Z)

'close Sap2000

SapObject.ApplicationExit False

Set SapModel = Nothing

Set SapObject = Nothing

End Sub

## Release Notes

Initial release in version 16.0.0.

## See Also

# GetLocalAxesAnalysys

## Syntax

SapObject.SapModel.SectCut.GetLocalAxesAnalysis

## VB6 Procedure

Function GetLocalAxesAnalysis(ByVal Name As String, ByRef Z As Double, ByRef Y As Double, ByRef X As Double, ByRef IsAdvanced As Boolean) As Long

## Parameters

Name

The name of an existing section cut.

Z

The Rotation about the Z axis.

Y

The rotation about the Y' axis where Y' is the orientation of the Y axis after rotation about the Z axis.

X

The rotation about the X'' axis where X'' is the orientation of the X axis after rotation about the Z axis and about the Y' axis.

IsAdvanced

Indicates if advanced local axes are specified.

## Remarks

This function gets the local axes angles for an existing section cut whose result type is Analysis.

The function returns zero if the angles are successfully retrieved, otherwise it returns a nonzero value.

## VBA Example

Sub GetSectionCutLocalAxesAnalysis()

'dimension variables

Dim SapObject As Sap2000v16.SapObject

Dim SapModel As cSapModel

Dim ret As Long

Dim Z As Double

Dim Y As Double

Dim X As Double

Dim IsAdvanced As Boolean

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

'set section cut local axes angles

ret = SapModel.SectCut.SetLocalAxesAnalysis ("SCut1", 10, 20, 30)

'get section cut local axes data

ret = SapModel.SectCut.GetLocalAxesAnalysis ("SCut1",Z, Y, X, IsAdvanced)

'close Sap2000

SapObject.ApplicationExit False

Set SapModel = Nothing

Set SapObject = Nothing

End Sub

## Release Notes

Initial release in version 16.0.0.

## See Also

[GetLocalAxesAdvancedAnalysis](GetLocalAxisAdvancedAnalysis.md)

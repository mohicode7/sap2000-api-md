# SetLocalAxisAdvancedAnalysis

## Syntax

SapObject.SapModel.SectCut.SetLocalAxesAdvancedAnalysis

## VB6 Procedure

Function SetLocalAxesAdvancedAnalysis(ByVal Name As String, ByVal Active As Boolean, ByVal AxVectOpt As Long, ByVal AxCSys As String, ByRef axdir() As Long, ByRef AxPt() As String, ByRef AxVect() As Double, ByVal Plane2 As Long, ByVal PlVectOpt As Long, ByVal PlCSys As String, ByRef PlDir() As Long, ByRef PlPt() As String, ByRef PlVect() As Double) As Long

## Parameters

Name

The name of an existing section cut.

Active

This is True if advanced local axes exist.

AxVectOpt, PlVectOpt

This is 1, 2 or 3, indicating the axis/plane reference vector option. This item applies only when the Active item is True.

1 = Coordinate direction

2 = Two joints

3 = User vector

AxCSys, PlCSys

The coordinate system used to define the axis/plane reference vector coordinate directions and the axis/plane user vector. This item applies when the Active item is True and the AxVectOpt/PlVectOpt item is 1 or 3.

AxDir, PlDir

This is an array dimensioned to 1 (2 integers), indicating the axis/plane reference vector primary and secondary coordinate directions, PlDir(0) and PlDir(1) respectively, taken at the object center in the specified coordinate system and used to determine the plane reference vector. This item applies when the Active item is True and the AxVectOpt/PlVectOpt item is 1. Possible coordinate direction values are:

1 = +X            -1 = -X

2 = +Y            -2 = -Y

3 = +Z            -3 = -Z

4 = +CR           -4 = -CR

5 = +CA           -5 = -CA

6 = +CZ           -6 = -CZ

7 = +SR           -7 = -SR

8 = +SA           -8 = -SA

9 = +SB           -9 = -SB

AxPt, PlPt

This is an array dimensioned to 1 (2 strings), indicating the labels of two joints that define the axis/plane reference vector. Either of these joints may be specified as None to indicate the center of the specified object.  If both joints are specified as None, they are not used to define the plane reference vector. This item applies when the Active item is True and the AxVectOpt/PlVectOpt item is 2.

AxVect, PlVect

This is an array dimensioned to 2 (3 doubles) that defines the axis/plane reference vector. This item applies when the Active item is True and the AxVectOpt/PlVectOpt item is 3.

Plane2

This is 12, 13, 21, 23, 31 or 32, indicating that the local plane determined by the plane reference vector is the 1-2, 1-3, 2-1, 2-3, 3-1, 0r 3-2 plane. This item applies only when the Active item is True.

## Remarks

This function sets the advanced local axes data for an existing section cut whose result type is Analysis.

The function returns zero if the data is successfully set, otherwise it returns a nonzero value.

## VBA Example

Sub SetSectionCutLocalAxesAdvancedAnalysis()

'dimension variables

Dim SapObject as cOAPI

Dim SapModel As cSapModel

Dim ret As Long

Dim MyAxDir(1) As Long

Dim MyAxPt(1) As String

Dim MyAxVect(2) As Double

Dim MyPlDir(1) As Long

Dim MyPlPt(1) As String

Dim MyPlVect(2) As Double

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

'define new group

ret = SapModel.GroupDef.SetGroup("Group1")

'add objects to group

ret = SapModel.PointObj.SetGroupAssign("1", "Group1")

ret = SapModel.FrameObj.SetGroupAssign("1", "Group1")

'define section cut

ret = SapModel.SectCut.SetByGroup("SCut1", "Group1", 1)

'set section cut advanced local axes data

**MyAxVect(0)=0.707**

MyAxVect(1)=0.707

MyAxVect(2)=0

MyPlDir(0) = 2

MyPlDir(1) = 3

ret = SapModel.SectCut.SetLocalAxesAdvancedAnalysis("SCut1", True, 3, "Global", MyAxDir, MyAxPt, MyAxVect, 12, 1, "Global", MyPlDir, MyPlPt, MyPlVect)

'close Sap2000

SapObject.ApplicationExit False

Set SapModel = Nothing

Set SapObject = Nothing

End Sub

## Release Notes

Initial release in version 16.0.0.

## See Also

SetLocalAxesAnalysis

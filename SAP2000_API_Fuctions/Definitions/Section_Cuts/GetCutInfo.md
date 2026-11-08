# GetCutInfo

## Syntax

SapObject.SapModel.SectCut.GetCutInfo

## VB6 Procedure

Function GetCutInfo(ByVal Name As String, ByRef GroupName As String, ByRef MyType As Long , ByRef Num As Long) As Long

## Parameters

Name

The name of an existing section cut.

GroupName

The name of the group associated with the section cut.

MyType

This is either 1, 2, 3 or 4 indicating the result type of the section cut.

1 = Analysis

2 = Design Wall

3 = Design Spandrel

4 = Design Slab

Num

The number of quadrilateral cutting planes defined for the section cut. If this number is zero then the section cut is defined by the associated group.

## Remarks

This function gets basic information about an existing section cut.

The function returns zero if the section cut information is successfully obtained, otherwise it returns a nonzero value.

## VBA Example

Sub GetSectionCutInfo()

'dimension variables

Dim SapObject As Sap2000v16.SapObject

Dim SapModel As cSapModel

Dim ret As Long

Dim GroupName As String

Dim MyType As Long

Dim Num As Long

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

'get basic section cut information

ret = SapModel.SectCut.GetCutInfo ("SCut1", GroupName, MyType, Num)

'close Sap2000

SapObject.ApplicationExit False

Set SapModel = Nothing

Set SapObject = Nothing

End Sub

## Release Notes

Initial release in version 16.0.0.

## See Also

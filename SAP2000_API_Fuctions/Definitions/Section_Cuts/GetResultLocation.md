# GetResultLocation

## Syntax

SapObject.SapModel.SectCut.GetResultLocation

## VB6 Procedure

Function GetResultLocation(ByVal Name As String, ByRef IsDefault As Boolean, ByRef X As Double, ByRef Y As Double, ByRef Z As Double) As Long

## Parameters

Name

The name of an existing section cut.

IsDefault

Indicates if the section cut results are reported at the default location.  If so, the X, Y and Z items are ignored.

X

The X coordinate of the section cut result location when it is not default

Y

The Y coordinate of the section cut result location when it is not default

Z

The Z coordinate of the section cut result location when it is not default

## Remarks

This function gets the results location for an existing section cut.

The function returns zero if the result location is successfully retrieved, otherwise it returns a nonzero value.

## VBA Example

Sub GetSectionCutResultsLocation()

'dimension variables

Dim SapObject As Sap2000v16.SapObject

Dim SapModel As cSapModel

Dim ret As Long

Dim IsDefault As Boolean

Dim X As Double

Dim Y As Double

Dim Z As Double

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

'set section cut result location

ret = SapModel.SectCut.SetResultLocation("SCut1", False, -288, 0, 10)

'get section cut result location

ret = SapModel.SectCut.GetResultLocation("SCut1", IsDefault, X, Y, Z)

'close Sap2000

SapObject.ApplicationExit False

Set SapModel = Nothing

Set SapObject = Nothing

End Sub

## Release Notes

Initial release in version 16.0.0.

## See Also

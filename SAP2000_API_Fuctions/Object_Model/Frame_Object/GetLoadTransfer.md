# GetLoadTransfer

## Syntax

SapObject.SapModel.FrameObj.GetLoadTransfer

## VB6 Procedure

Function GetLoadTransfer(ByVal Name As String, ByRef Val As Boolean) As Long

## Parameters

Name

The name of an existing frame.

Val

This boolean value indicates if load is allowed to be transferred from area objects to this frame object.

## Remarks

This function returns the load transfer option for a frame object.  It indicates whether the frame receives load from an area object when the area object is loaded with a load of type uniform to frame.

The function returns zero if the load transfer option is successfully returned, otherwise it returns a nonzero value.

## VBA Example

Sub GetLoadTransferOption()

'dimension variables

Dim SapObject as cOAPI

Dim SapModel As cSapModel

Dim ret As Long

Dim val As Boolean

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

'get the load transfer option to False for frame object 1

ret = SapModel.FrameObj.GetLoadTransfer("1", val)

'close Sap2000

SapObject.ApplicationExit False

Set SapModel = Nothing

Set SapObject = Nothing

End Sub

## Release Notes

## Initial release in version 16.0.0.

## See Also

## [SetLoadTransfer](SetLoadTransfer.md)

##

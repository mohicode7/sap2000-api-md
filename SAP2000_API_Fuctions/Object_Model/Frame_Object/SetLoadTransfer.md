# SetLoadTransfer

## Syntax

SapObject.SapModel.FrameObj.SetLoadTransfer

## VB6 Procedure

Function SetLoadTransfer(ByVal Name As String, ByVal Val As Boolean, Optional ByVal ItemType As eItemType = Object) As Long

## Parameters

Name

The name of an existing frame object or group, depending on the value of the ItemType item.

Val

This boolean value indicates if load is allowed to be transferred from area objects to this frame object.

ItemType

This is one of the following items in the eItemType enumeration:

Object = 0

Group = 1

SelectedObjects = 2

If this item is Object, the assignment is made to the frame object specified by the Name item.

If this item is Group, the assignment is made to all frame objects in the group specified by the Name item.

If this item is SelectedObjects, assignment is made to all selected frame objects, and the Name item is ignored.

## Remarks

This function returns the load transfer option for frame objects.  It indicates whether the frame receives load from an area object when the area object is loaded with a load of type uniform to frame.

The function returns zero if the load transfer option is successfully returned, otherwise it returns a nonzero value.

## VBA Example

Sub GetLoadTransferOption()

'dimension variables

Dim SapObject as cOAPI

Dim SapModel As cSapModel

Dim ret As Long

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

'set the load transfer option to False for frame object 1

ret = SapModel.FrameObj.SetLoadTransfer("1", False)

'close Sap2000

SapObject.ApplicationExit False

Set SapModel = Nothing

Set SapObject = Nothing

End Sub

## Release Notes

## Initial release in version 16.0.0.

## See Also

## [GetLoadTransfer](GetLoadTransfer.md)

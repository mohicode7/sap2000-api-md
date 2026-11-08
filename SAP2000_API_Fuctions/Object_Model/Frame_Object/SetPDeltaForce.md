# SetPDeltaForce

## Syntax

SapObject.SapModel.FrameObj.SetPDeltaForce

## VB6 Procedure

Function SetPDeltaForce(ByVal Name As String, ByVal PDeltaForce As Double, ByVal Dir As Long, ByVal Replace As Boolean, Optional ByVal CSys As String = "Global", Optional ByVal ItemType As eItemType = Object) As Long

## Parameters

Name

The name of an existing frame object or group, depending on the value of the ItemType item.

PDeltaForce

The P-Delta force assigned to the frame object. [F]

Dir

This is 0, 1, 2 or 3, indicating the direction of the P-Delta force assignment.

0 = Frame object local 1-axis direction

1 = Projected X direction in CSys coordinate system

2 = Projected Y direction in CSys coordinate system

3 = Projected Z direction in CSys coordinate system

Replace

If this item is True, all existing P-Delta force assignments to the frame object are removed before assigning the specified P-Delta force. If it is False, the specified P-Delta force is added to any existing P-Delta forces already assigned to the frame object.

CSys

This is the name of the coordinate system in which the projected X, Y or Z direction P-Delta forces are defined. This item does not apply if the Dir item is zero (frame object local 1-axis direction).

ItemType

This is one of the following items in the eItemType enumeration:

Object = 0

Group = 1

SelectedObjects = 2

If this item is Object, the assignment is made to the frame object specified by the Name item.

If this item is Group, the assignment is made to all frame objects in the group specified by the Name item.

If this item is SelectedObjects, assignment is made to all selected frame objects, and the Name item is ignored.

## Remarks

This function assigns P-Delta forces to straight frame objects. P-Delta force assignments do not apply to curved frames.

The function returns zero if the assignments are successfully applied, otherwise it returns a nonzero value.

## VBA Example



```
Sub AssignFramePDeltaForce()

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

      ret = SapModel.File.New2DFrame(PortalFrame, 3, 124, 3, 200)



   'assign P-Delta force

      ret = SapModel.FrameObj.SetPDeltaForce("ALL", 100, 0, True, , Group)



   'close Sap2000

      SapObject.ApplicationExit False

      Set SapModel = Nothing

      Set SapObject = Nothing

End Sub
```



## Release Notes

Initial release in version 11.00.

## See Also

[GetPDeltaForce](GetPDeltaForce_%7BFrame_Object%7D.md)

[DeletePDeltaForce](DeletePDeltaForce.md)
